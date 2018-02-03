<?php

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleExcel;

use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class SimpleExcel
{
    private $contentTypes = [
        'csv' => 'text/csv',
        'xls' => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    private $writers = [
        'csv' => 'Csv',
        'xls' => 'Xls',
        'xlsx' => 'Xlsx',
    ];

    private $sheets;

    public function __construct()
    {
        $this->sheets = [];
    }

    public function loadFromArray(array $data): void
    {
        // If the data is not multidimensional make it so
        if (!is_array(current($data))) {
            $data = [$data];
        }

        foreach ($data as $sheetName => $sheet) {
            $this->sheets[$sheetName] = $sheet;
        }
    }

    public function loadFromFile(string $filename, array $sheetNames = []): void
    {
        $excel = IOFactory::load($filename);

        $this->loadFromExcel($excel, $sheetNames);
    }

    public function loadFromExcel(Spreadsheet $excel, array $sheetNames = []): void
    {
        foreach ($excel->getWorksheetIterator() as $sheet) {
            if (0 == count($sheetNames) || in_array($sheet->getTitle(), $sheetNames)) {
                $this->loadFromSheet($sheet);
            }
        }
    }

    public function loadFromSheet(Worksheet $excelSheet): void
    {
        $sheet = [];

        $sheetData = $excelSheet->toArray('', false, false);

        // Remove possible empty leading rows
        while ($sheetData[0][0] == '' && count($sheetData) > 0) {
            array_shift($sheetData);
        }

        // First row always contains the headers
        $columns = array_shift($sheetData);

        // Remove headers from the end until first name is found
        for ($i = count($columns) - 1; $i >= 0; --$i) {
            if ('' == $columns[$i]) {
                unset($columns[$i]);
            } else {
                break;
            }
        }

        // Next rows contain the actual data
        foreach ($sheetData as $row) {
            // Ignore empty rows
            if ('' == trim(implode('', $row))) {
                continue;
            }

            // Map data to column names
            $associativeRow = [];
            foreach ($row as $key => $value) {
                if (!isset($columns[$key])) {
                    continue;
                }
                if ('null' == strtolower($value)) {
                    $value = null;
                }
                $associativeRow[$columns[$key]] = $value;
            }
            $sheet[] = $associativeRow;
        }

        $this->sheets[$excelSheet->getTitle()] = $sheet;
    }

    public function saveToArray(array $sheetNames = []): array
    {
        $sheets = [];

        foreach ($this->sheets as $sheetName => $sheet) {
            if (0 == count($sheetNames) || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheet;
            }
        }

        return $sheets;
    }

    public function saveToExcel(array $sheetNames = []): Spreadsheet
    {
        $excel = new Spreadsheet();
        $excel->removeSheetByIndex(0);

        foreach ($this->sheets as $sheetName => $sheet) {
            // Only process requested sheets
            if (count($sheetNames) > 0 && !in_array($sheetName, $sheetNames)) {
                continue;
            }

            $excelSheet = $excel->createSheet();
            $excelSheet->setTitle($sheetName);

            // When no content is available leave sheet empty
            if (0 == count($sheet)) {
                continue;
            }

            // Set column headers
            $headers = array_keys($sheet[0]);
            array_unshift($sheet, $headers);

            // Place values in sheet
            $rowId = 1;
            foreach ($sheet as $row) {
                $colId = ord('A');
                foreach ($row as $value) {
                    if (null === $value) {
                        $value = 'NULL';
                    }
                    $excelSheet->setCellValue(chr($colId).$rowId, $value);
                    ++$colId;
                }
                ++$rowId;
            }
        }

        return $excel;
    }

    public function saveToFile(string $filename, array $sheetNames = []): void
    {
        $writer = $this->getWriterByFilename($filename, $sheetNames);
        $writer->save($filename);
    }

    public function saveToOutput(string $filename, array $sheetNames = [], bool $setHeaders = true): void
    {
        if ($setHeaders) {
            $headers = $this->getHeadersByFilename($filename);
            foreach ($headers as $key => $value) {
                header($key.': '.$value);
            }
        }

        $writer = $this->getWriterByFilename($filename, $sheetNames);
        $writer->save('php://output');
    }

    public function saveToString(string $filename, array $sheetNames = []): string
    {
        ob_start();

        $this->saveToOutput($filename, $sheetNames, false);

        return ob_get_clean();
    }

    protected function getContentTypeByFilename($filename): string
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->contentTypes[$extension])) {
            throw new Exception(sprintf('No content type defined for file extension "%s"', $extension));
        }

        return $this->contentTypes[$extension];
    }

    public function getWriterByFilename(string $filename, array $sheetNames = []): IWriter
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->writers[$extension])) {
            throw new Exception(sprintf('No writer defined for file extension "%s"', $extension));
        }

        $excel = $this->saveToExcel($sheetNames);

        return IOFactory::createWriter($excel, $this->writers[$extension]);
    }

    public function getHeadersByFilename(string $filename): array
    {
        $headers = [
            'Content-Disposition' => 'attachment; filename='.$filename,
            'Cache-Control' => 'max-age=0',
            'Content-Type' => $this->getContentTypeByFilename($filename).'; charset=utf-8',
        ];

        return $headers;
    }

    private function getExtension(string $filename): string
    {
        return strtolower(substr(strrchr($filename, '.'), 1));
    }
}
