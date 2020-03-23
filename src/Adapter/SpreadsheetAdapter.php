<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SpreadsheetAdapter extends AbstractAdapter
{
    public function load($data, array $sheetNames = null): array
    {
        if (!$data instanceof Spreadsheet) {
            throw new SimpleSpreadsheetException('Invalid spreadsheet data');
        }

        $sheets = [];

        foreach ($data->getWorksheetIterator() as $sheet) {
            if (null === $sheetNames || in_array($sheet->getTitle(), $sheetNames)) {
                $sheetData = $sheet->toArray('', false, false);

                // Remove possible empty leading rows
                while ('' == $sheetData[0][0] && count($sheetData) > 0) {
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
                        if ('null' == strtolower(strval($value))) {
                            $value = null;
                        }
                        $associativeRow[$columns[$key]] = $value;
                    }
                    $sheets[$sheet->getTitle()][] = $associativeRow;
                }
            }
        }

        return $sheets;
    }

    public function save(array $data, array $sheetNames = null, array $options = [])
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->removeSheetByIndex(0);

        foreach ($data as $sheetName => $sheetData) {
            // Only process requested sheets
            if (null !== $sheetNames && !in_array($sheetName, $sheetNames)) {
                continue;
            }

            $sheet = $spreadsheet->createSheet();
            $sheet->setTitle($sheetName);

            // When no content is available leave sheet empty
            if (0 == count($sheetData)) {
                continue;
            }

            // Set column headers
            $headers = array_keys(current($sheetData));
            array_unshift($sheetData, $headers);

            // Place values in sheet
            $rowId = 1;
            foreach ($sheetData as $row) {
                $colId = ord('A');
                foreach ($row as $value) {
                    if (null === $value) {
                        $value = 'NULL';
                    }
                    $sheet->setCellValue(chr($colId).$rowId, $value);
                    ++$colId;
                }
                ++$rowId;
            }
        }

        return $spreadsheet;
    }

    public function supports($data): bool
    {
        return $data instanceof Spreadsheet;
    }
}
