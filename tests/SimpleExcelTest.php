<?php

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleExcel\SimpleExcelTest;

use Endroid\SimpleExcel\SimpleExcel;
use PHPUnit_Framework_TestCase;

class SimpleExcelTest extends PHPUnit_Framework_TestCase
{
    /**
     * Tests load and save.
     */
    public function testLoadAndSave()
    {
        $excel = new SimpleExcel();
        $excel->loadFromFile(__DIR__.'/data/data.xlsx');
        $excel->loadFromArray(array(
            'Sheet A' => array(
                array('col1' => 'a', 'col2' => 'b', 'col3' => 'c'),
                array('col1' => 'b', 'col2' => 'c', 'col3' => 'd')
            )
        ));

        $data = $excel->saveToArray();

        $this->assertTrue(count($data) == 3);
        $this->assertTrue(count($data['Sheet A']) == 2);
        $this->assertNull($data['sheet1'][1]['col2']);
    }
}
