<?php

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleExcel;

use Exception;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Worksheet;
use PhpOffice\PhpWord\Writer\WriterInterface;

class SimpleExcel
{
    /**
     * @var array
     */
    protected $contentTypes = [
        'csv' => 'text/csv',
        'xls' => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    /**
     * @var array
     */
    protected $writers = [
        'csv' => 'CSV',
        'xls' => 'Excel5',
        'xlsx' => 'Excel2007',
    ];

    /**
     * @var array
     */
    protected $sheets;

    /**
     * Creates a new instance.
     */
    public function __construct()
    {
        $this->sheets = [];
    }

    /**
     * Loads sheets from an array.
     *
     * @param array $data
     */
    public function loadFromArray(array $data)
    {
        // If the data is not multidimensional make it so
        if (!is_array(current($data))) {
            $data = [$data];
        }

        foreach ($data as $sheetName => $sheet) {
            $this->sheets[$sheetName] = $sheet;
        }
    }

    /**
     * Loads sheets from a file.
     *
     * @param string $filename
     * @param array  $sheetNames
     */
    public function loadFromFile($filename, array $sheetNames = [])
    {
        $excel = PHPExcel_IOFactory::load($filename);

        $this->loadFromExcel($excel, $sheetNames);
    }

    /**
     * Loads sheets from an Excel document.
     *
     * @param PHPExcel $excel
     * @param array    $sheetNames
     */
    public function loadFromExcel(PHPExcel $excel, array $sheetNames = [])
    {
        foreach ($excel->getWorksheetIterator() as $sheet) {
            if (0 == count($sheetNames) || in_array($sheet->getTitle(), $sheetNames)) {
                $this->loadFromSheet($sheet);
            }
        }
    }

    /**
     * Loads an Excel document sheet.
     *
     * @param PHPExcel_Worksheet $excelSheet
     */
    public function loadFromSheet(PHPExcel_Worksheet $excelSheet)
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

    /**
     * Saves to an array.
     *
     * @param array $sheetNames
     *
     * @return array
     */
    public function saveToArray(array $sheetNames = [])
    {
        $sheets = [];

        foreach ($this->sheets as $sheetName => $sheet) {
            if (0 == count($sheetNames) || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheet;
            }
        }

        return $sheets;
    }

    /**
     * Saves to an Excel document.
     *
     * @param array $sheetNames
     *
     * @return PHPExcel
     */
    public function saveToExcel(array $sheetNames = [])
    {
        $excel = new PHPExcel();
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

    /**
     * Saves to a file.
     *
     * @param string $filename
     * @param array  $sheetNames
     */
    public function saveToFile($filename, array $sheetNames = [])
    {
        $writer = $this->getWriterByFilename($filename, $sheetNames);
        $writer->save($filename);
    }

    /**
     * Saves to output.
     *
     * @param $filename
     * @param array $sheetNames
     * @param bool  $setHeaders
     *
     * @throws Exception
     */
    public function saveToOutput($filename, array $sheetNames = [], $setHeaders = true)
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

    /**
     * Saves to a string.
     *
     * @param $filename
     * @param array $sheetNames
     *
     * @return string
     */
    public function saveToString($filename, array $sheetNames = [])
    {
        ob_start();

        $this->saveToOutput($filename, $sheetNames, false);

        return ob_get_clean();
    }

    /**
     * Returns the content type for a specific file name.
     *
     * @param $filename
     *
     * @return string
     *
     * @throws Exception
     */
    protected function getContentTypeByFilename($filename)
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->contentTypes[$extension])) {
            throw new Exception(sprintf('No content type defined for file extension "%s"', $extension));
        }

        return $this->contentTypes[$extension];
    }

    /**
     * Returns the writer for a specific file name.
     *
     * @param string $filename
     * @param array  $sheetNames
     *
     * @return WriterInterface
     *
     * @throws Exception
     */
    public function getWriterByFilename($filename, array $sheetNames = [])
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->writers[$extension])) {
            throw new Exception(sprintf('No writer defined for file extension "%s"', $extension));
        }

        $excel = $this->saveToExcel($sheetNames);

        return PHPExcel_IOFactory::createWriter($excel, $this->writers[$extension]);
    }

    /**
     * Returns the headers for a specific file name.
     *
     * @param $filename
     *
     * @return array
     */
    public function getHeadersByFilename($filename)
    {
        $headers = [
            'Content-Disposition' => 'attachment; filename='.$filename,
            'Cache-Control' => 'max-age=0',
            'Content-Type' => $this->getContentTypeByFilename($filename).'; charset=utf-8',
        ];

        return $headers;
    }

    /**
     * Returns the extension of a file name.
     *
     * @param $filename
     *
     * @return string
     */
    protected function getExtension($filename)
    {
        return strtolower(substr(strrchr($filename, '.'), 1));
    }
}
