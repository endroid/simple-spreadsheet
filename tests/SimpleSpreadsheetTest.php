<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\SimpleSpreadsheetTest;

use Endroid\SimpleSpreadsheet\Adapter\ArrayAdapter;
use Endroid\SimpleSpreadsheet\Adapter\FileAdapter;
use Endroid\SimpleSpreadsheet\Adapter\JsonAdapter;
use Endroid\SimpleSpreadsheet\Adapter\ResponseAdapter;
use Endroid\SimpleSpreadsheet\Adapter\SpreadsheetAdapter;
use Endroid\SimpleSpreadsheet\SimpleSpreadsheet;
use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\Attributes\TestDox;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use Symfony\Component\HttpFoundation\Response;

final class SimpleSpreadsheetTest extends TestCase
{
    public function testMalformedSheetArraysAreRejected(): void
    {
        $adapter = new ArrayAdapter();
        $this->assertFalse($adapter->supports(['sheet' => null]));
        $this->assertFalse($adapter->supports(['sheet' => 'not an array']));
        $this->assertFalse((new JsonAdapter())->supports('{"sheet":123}'));
        $this->assertTrue((new JsonAdapter())->supports('{"sheet":[]}'));
        $this->expectException(\InvalidArgumentException::class);
        $adapter->load(['sheet' => 123]);
    }

    public function testMalformedJsonSheetsAreRejected(): void
    {
        $this->expectException(\InvalidArgumentException::class);
        (new JsonAdapter())->load('{"sheet":123}');
    }

    /** @return iterable<string, array{string, array<mixed>}> */
    public static function invalidOutputOptions(): iterable
    {
        foreach ([[], ['value' => null], ['value' => 123], ['value' => []]] as $index => $options) {
            yield 'file '.$index => [FileAdapter::class, array_key_exists('value', $options) ? ['path' => $options['value']] : []];
            yield 'response '.$index => [ResponseAdapter::class, array_key_exists('value', $options) ? ['filename' => $options['value']] : []];
        }
        yield 'unsupported response extension' => [ResponseAdapter::class, ['filename' => 'sheet.html']];
    }

    /** @param array<mixed> $options */
    #[DataProvider('invalidOutputOptions')]
    public function testInvalidOutputOptionsAreRejected(string $adapter, array $options): void
    {
        $this->expectException(SimpleSpreadsheetException::class);
        (new SimpleSpreadsheet())->save($adapter, options: $options);
    }

    public function testMalformedRowsAreRejected(): void
    {
        $this->expectException(SimpleSpreadsheetException::class);
        (new SpreadsheetAdapter())->save(['sheet' => [['header' => 'valid'], 'not a row']]);
    }

    public function testEmptyWorksheetLoadsAsEmptySheet(): void
    {
        $source = new Spreadsheet();
        $source->getActiveSheet()->setTitle('empty');
        $this->assertSame(['empty' => []], (new SpreadsheetAdapter())->load($source));
    }

    /** @return iterable<string, array{string}> */
    public static function drawingCells(): iterable
    {
        yield 'header' => ['A1'];
        yield 'data' => ['A2'];
    }

    #[DataProvider('drawingCells')]
    public function testNonScalarWorksheetCellsAreRejected(string $coordinate): void
    {
        $source = new Spreadsheet();
        $sheet = $source->getActiveSheet();
        $sheet->setCellValue('A1', 'header');
        $sheet->setCellValue('A2', 'value');
        $sheet->getCell($coordinate)->setValueExplicit(
            new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing(),
            \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_DRAWING_IN_CELL
        );

        $this->expectException(SimpleSpreadsheetException::class);
        (new SpreadsheetAdapter())->load($source);
    }

    public function testUnknownAdapterIsRejected(): void
    {
        $this->expectException(SimpleSpreadsheetException::class);
        (new SimpleSpreadsheet())->save('MissingAdapter');
    }

    public function testNumericSheetNamesAreSelectedExactly(): void
    {
        $data = ['0' => [['value' => 'zero']], '00' => [['value' => 'double zero']]];
        foreach (['0', '00'] as $name) {
            $spreadsheet = new SimpleSpreadsheet();
            $spreadsheet->load($data, [$name]);
            $this->assertSame([$name => $data[$name]], $spreadsheet->save(ArrayAdapter::class));

            $spreadsheet = new SimpleSpreadsheet();
            $spreadsheet->load($data);
            $this->assertSame([$name => $data[$name]], $spreadsheet->save(ArrayAdapter::class, [$name]));
        }
    }

    public function testSpreadsheetSheetNamesAreSelectedExactly(): void
    {
        $source = new Spreadsheet();
        $source->getActiveSheet()->setTitle('0')->fromArray([['value'], ['zero']]);
        $source->createSheet()->setTitle('00')->fromArray([['value'], ['double zero']]);
        $adapter = new SpreadsheetAdapter();

        $this->assertSame([0 => [['value' => 'zero']]], $adapter->load($source, ['0']));
        $this->assertSame(['00' => [['value' => 'double zero']]], $adapter->load($source, ['00']));
        $this->assertSame(['00'], $adapter->save(['00' => [['value' => 'double zero']]], ['00'])->getSheetNames());
        $this->assertSame([], $adapter->save(['00' => [['value' => 'double zero']]], ['0'])->getSheetNames());
        $this->assertSame(['0'], $adapter->save([0 => [['value' => 'zero']]], ['0'])->getSheetNames());
        $this->assertSame([], $adapter->save([0 => [['value' => 'zero']]], ['00'])->getSheetNames());
    }

    public function testFalseHeadersAreNotEmptyHeaders(): void
    {
        $source = new Spreadsheet();
        $sheet = $source->getActiveSheet();
        $sheet->setTitle('headers');
        $sheet->setCellValue('A1', false);
        $sheet->setCellValue('B1', 'last');
        $sheet->setCellValue('A2', 'first value');
        $sheet->setCellValue('B2', 'last value');

        $this->assertSame(
            ['headers' => [[0 => 'first value', 'last' => 'last value']]],
            (new SpreadsheetAdapter())->load($source)
        );

        $sheet->setCellValue('A1', 'first');
        $sheet->setCellValue('B1', false);
        $this->assertSame(
            ['headers' => [['first' => 'first value', 0 => 'last value']]],
            (new SpreadsheetAdapter())->load($source)
        );
    }

    #[TestDox('Append sheets and rows')]
    public function testAppend(): void
    {
        $spreadsheet = new SimpleSpreadsheet();
        $spreadsheet->load(__DIR__.'/data/data.xlsx');

        $spreadsheet->load([
            'sheet1' => [
                ['col1' => 'a', 'col2' => 'b', 'col3' => 'c'],
                ['col1' => 'b', 'col2' => 'c', 'col3' => 'd'],
            ],
            'sheet3' => [
                ['col1' => 'a', 'col2' => 'b', 'col3' => 'c'],
                ['col1' => 'b', 'col2' => 'c', 'col3' => 'd'],
            ],
        ]);

        $data = $spreadsheet->save(ArrayAdapter::class);

        $this->assertEquals(3, count($data));
        $this->assertEquals(4, count($data['sheet1']));
        $this->assertNull($data['sheet1'][1]['col2']);
    }

    #[TestDox('Save to file')]
    public function testSaveToFile(): void
    {
        $spreadsheet = new SimpleSpreadsheet();
        $spreadsheet->load(__DIR__.'/data/data.xlsx');

        if (!is_dir(__DIR__.'/output')) {
            mkdir(__DIR__.'/output');
        }

        $targetPath = __DIR__.'/output/data.xlsx';
        $spreadsheet->save(
            FileAdapter::class,
            ['sheet1'],
            ['path' => $targetPath]
        );

        $this->assertFileExists($targetPath);
    }

    #[TestDox('Save to response object')]
    public function testSaveToResponse(): void
    {
        $spreadsheet = new SimpleSpreadsheet();
        $spreadsheet->load(__DIR__.'/data/data.xlsx');

        $response = $spreadsheet->save(
            ResponseAdapter::class,
            ['sheet1'],
            ['filename' => 'data.xlsx']
        );

        $this->assertInstanceOf(Response::class, $response);
    }
}
