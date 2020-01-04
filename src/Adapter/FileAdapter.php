<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

use PhpOffice\PhpSpreadsheet\IOFactory;

class FileAdapter extends AbstractAdapter
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

    public function load($data, array $sheetNames = null): array
    {
        $excel = IOFactory::load($data);
    }

    public function save(array $data, array $sheetNames = null)
    {
        $writer = $this->getWriterByFilename($filename, $sheetNames);
        $writer->save($filename);
    }

    public function supports($data): bool
    {
        if (!is_string($data)) {
            return false;
        }

        return file_exists($data);
    }
}
