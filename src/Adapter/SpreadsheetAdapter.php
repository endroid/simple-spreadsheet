<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SpreadsheetAdapter extends AbstractAdapter
{
    public function load($data, array $sheetNames = null): array
    {
    }

    public function save(array $data, array $sheetNames = null)
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
            $headers = array_keys(current($sheet));
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

    public function supports($data): bool
    {
        return $data instanceof Spreadsheet;
    }
}
