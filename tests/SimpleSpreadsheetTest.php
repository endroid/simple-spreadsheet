<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\SimpleSpreadsheetTest;

use Endroid\SimpleSpreadsheet\Adapter\ArrayAdapter;
use Endroid\SimpleSpreadsheet\Adapter\FileAdapter;
use Endroid\SimpleSpreadsheet\Adapter\ResponseAdapter;
use Endroid\SimpleSpreadsheet\SimpleSpreadsheet;
use PHPUnit\Framework\TestCase;
use Symfony\Component\HttpFoundation\Response;

class SimpleSpreadsheetTest extends TestCase
{
    /**
     * @testdox Append sheets and rows
     */
    public function testAppend()
    {
        $spreadsheet = $this->loadSpreadsheet();

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

    /**
     * @testdox Save to file
     */
    public function testSaveToFile(): void
    {
        $spreadsheet = $this->loadSpreadsheet();

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

    /**
     * @testdox Save to response object
     */
    public function testSaveToResponse(): void
    {
        $spreadsheet = $this->loadSpreadsheet();

        $response = $spreadsheet->save(
            ResponseAdapter::class,
            ['sheet1'],
            ['filename' => 'data.xlsx']
        );

        $this->assertInstanceOf(Response::class, $response);
    }

    private function loadSpreadsheet(): SimpleSpreadsheet
    {
        $spreadsheet = new SimpleSpreadsheet();
        $spreadsheet->load(__DIR__.'/data/data.xlsx');

        return $spreadsheet;
    }
}
