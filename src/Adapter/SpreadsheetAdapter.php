<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * @implements AdapterInterface<Spreadsheet, Spreadsheet>
 */
final readonly class SpreadsheetAdapter implements AdapterInterface
{
    #[\Override]
    public function supports(mixed $data): bool
    {
        return $data instanceof Spreadsheet;
    }

    /** @param mixed $data */
    #[\Override]
    public function load(mixed $data, ?array $sheetNames = null): array
    {
        if (!$data instanceof Spreadsheet) {
            throw new SimpleSpreadsheetException('Invalid spreadsheet data');
        }

        $sheets = [];

        foreach ($data->getWorksheetIterator() as $sheet) {
            if (!(null === $sheetNames || in_array($sheet->getTitle(), $sheetNames, strict: true))) {
                continue;
            }

            // Default numeric row/column indices form a list of lists.
            /** @var list<list<mixed>> $sheetData */
            $sheetData = $sheet->toArray('', false, false);

            // Remove possible empty leading rows
            while ('' === ($sheetData[0][0] ?? null)) {
                array_shift($sheetData);
            }

            // The first row always contains the headers
            $columns = array_map($this->requireCell(...), array_shift($sheetData) ?? []);

            // Remove headers from the end until first name is found
            for ($i = count($columns) - 1; $i >= 0; --$i) {
                if ('' !== ($columns[$i] ?? null)) {
                    break;
                }
                unset($columns[$i]);
            }

            // Headers are processed: now create the sheet array
            $sheets[$sheet->getTitle()] = [];

            // The following rows contain the actual data
            foreach ($sheetData as $row) {
                $row = array_map($this->requireCell(...), $row);
                // Ignore empty rows
                if ('' === trim(implode('', $row))) {
                    continue;
                }

                $sheets[$sheet->getTitle()][] = $this->mapRow($columns, $row);
            }
        }

        return $sheets;
    }

    /**
     * @param array<int, scalar> $columns
     * @param list<scalar> $row
     * @return array<array-key, scalar|null>
     */
    private function mapRow(array $columns, array $row): array
    {
        $associativeRow = [];
        foreach ($row as $key => $value) {
            if (!array_key_exists($key, $columns)) {
                continue;
            }
            if ('null' === strtolower(strval($value))) {
                $value = null;
            }
            $column = $columns[$key];
            $column = is_string($column) ? $column : (int) $column;
            $associativeRow[$column] = $value;
        }

        return $associativeRow;
    }

    #[\Override]
    public function save(array $data, ?array $sheetNames = null, array $options = []): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->removeSheetByIndex(0);

        foreach ($data as $sheetName => $sheetData) {
            // Only process requested sheets
            if (null !== $sheetNames && !in_array((string) $sheetName, $sheetNames, strict: true)) {
                continue;
            }

            $sheet = $spreadsheet->createSheet();
            $sheet->setTitle((string) $sheetName);

            // When no content is available leave sheet empty
            if (0 === count($sheetData)) {
                continue;
            }

            // Set column headers
            $sheetData = array_map($this->requireRow(...), $sheetData);
            $headers = array_keys($this->requireRow(current($sheetData)));
            array_unshift($sheetData, $headers);

            // Place values in sheet
            $rowId = 1;
            foreach ($sheetData as $row) {
                $colId = ord('A');
                foreach (array_keys($row) as $column) {
                    $sheet->setCellValue(chr($colId) . $rowId, $row[$column] ?? 'NULL');
                    ++$colId;
                }
                ++$rowId;
            }
        }

        return $spreadsheet;
    }

    private function requireCell(mixed $value): string|int|float|bool
    {
        if (!is_scalar($value)) {
            throw new SimpleSpreadsheetException('Only scalar worksheet cell values are supported');
        }

        return $value;
    }

    /** @return array<mixed> */
    private function requireRow(mixed $row): array
    {
        if (!is_array($row)) {
            throw new SimpleSpreadsheetException('Each sheet row must be an array');
        }

        return $row;
    }
}
