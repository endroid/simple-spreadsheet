<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\SimpleSpreadsheetTest;

use Endroid\SimpleSpreadsheet\Adapter\ArrayAdapter;
use Endroid\SimpleSpreadsheet\SimpleSpreadsheet;
use PHPUnit\Framework\TestCase;

class SimpleSpreadsheetTest extends TestCase
{
    /**
     * @testdox Test load and save
     */
    public function testLoadAndSave()
    {
        $spreadsheet = new SimpleSpreadsheet();
        $spreadsheet->load(__DIR__.'/data/data.xlsx');
        $spreadsheet->load([
            'Sheet A' => [
                ['col1' => 'a', 'col2' => 'b', 'col3' => 'c'],
                ['col1' => 'b', 'col2' => 'c', 'col3' => 'd'],
            ],
        ]);

        $data = $spreadsheet->save(ArrayAdapter::class);

        $this->assertTrue(3 == count($data));
        $this->assertTrue(2 == count($data['Sheet A']));
        $this->assertNull($data['sheet1'][1]['col2']);
    }
}
