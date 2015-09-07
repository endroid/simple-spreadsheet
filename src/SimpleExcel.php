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
    protected $contentTypes = array(
        'csv' => 'text/csv',
        'xls' => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    /**
     * @var array
     */
    protected $writers = array(
        'csv' => 'CSV',
        'xls' => 'Excel5',
        'xlsx' => 'Excel2007',
    );

    /**
     * @var array
     */
    protected $sheets;

    /**
     * Creates a new instance.
     */
    public function __construct()
    {
        $this->sheets = array();
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
            $data = array($data);
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
    public function loadFromFile($filename, array $sheetNames = array())
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
    public function loadFromExcel(PHPExcel $excel, array $sheetNames = array())
    {
        foreach ($excel->getWorksheetIterator() as $sheet) {
            if (count($sheetNames) == 0 || in_array($sheet->getTitle(), $sheetNames)) {
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
        $sheet = array();

        $sheetData = $excelSheet->toArray('', false, false);

        // Remove possible empty leading rows
        while ($sheetData[0][0] == '' && count($sheetData) > 0) {
            array_shift($sheetData);
        }

        // First row always contains the headers
        $columns = array_shift($sheetData);

        // Remove headers from the end until first name is found
        for ($i = count($columns) - 1; $i >= 0; --$i) {
            if ($columns[$i] == '') {
                unset($columns[$i]);
            } else {
                break;
            }
        }

        // Next rows contain the actual data
        foreach ($sheetData as $row) {

            // Ignore empty rows
            if (trim(implode('', $row)) == '') {
                continue;
            }

            // Map data to column names
            $associativeRow = array();
            foreach ($row as $key => $value) {
                if (!isset($columns[$key])) {
                    continue;
                }
                if (strtolower($value) == 'null') {
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
    public function saveToArray(array $sheetNames = array())
    {
        $sheets = array();

        foreach ($this->sheets as $sheetName => $sheet) {
            if (count($sheetNames) == 0 || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheet;
            }
        }

        return $sheets;
    }

    /**
     * Saves to a file.
     *
     * @param string $filename
     * @param array  $sheetNames
     */
    public function saveToFile($filename, array $sheetNames = array())
    {
        $excel = $this->saveToExcel($sheetNames);

        $writer = $this->getWriterByFilename($excel, $filename);
        $writer->save($filename);
    }

    /**
     * Saves to an Excel document.
     *
     * @param array $sheetNames
     *
     * @return PHPExcel
     */
    public function saveToExcel(array $sheetNames = array())
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
            if (count($sheet) == 0) {
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
                    if ($value === null) {
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
     * Outputs the file.
     *
     * @param string $filename
     * @param array  $sheetNames
     */
    public function output($filename, array $sheetNames = array())
    {
        $excel = $this->saveToExcel($sheetNames);

        ob_end_clean();

        $this->setHeadersForFilename($filename);

        $writer = $this->getWriterByFilename($excel, $filename);
        $writer->save('php://output');
        die;
    }

    /**
     * Sets the response headers for the given file name.
     *
     * @param string $filename
     */
    protected function setHeadersForFilename($filename)
    {
        header('Content-Disposition: attachment; filename='.$filename);
        header('Cache-Control: max-age=0');
        header('Content-type: '.$this->getContentTypeByFilename($filename).'; charset=utf-8');
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
     * @param PHPExcel $excel
     * @param string   $filename
     *
     * @return WriterInterface
     *
     * @throws Exception
     */
    protected function getWriterByFilename($excel, $filename)
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->writers[$extension])) {
            throw new Exception(sprintf('No writer defined for file extension "%s"', $extension));
        }

        return PHPExcel_IOFactory::createWriter($excel, $this->writers[$extension]);
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
