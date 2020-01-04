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
use PhpOffice\PhpSpreadsheet\IOFactory;

class FileAdapter extends SpreadsheetAdapter
{
    private $contentTypes = [
        'csv' => 'text/csv',
        'xls' => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    public function load($data, array $sheetNames = null): array
    {
        $spreadsheet = IOFactory::load($data);

        return parent::load($spreadsheet, $sheetNames);
    }

    public function save(array $data, array $sheetNames = null, array $options = [])
    {
        if (!isset($options['path'])) {
            throw new SimpleSpreadsheetException('Please specify the output path via options');
        }

        $path = $options['path'];
        $spreadsheet = parent::save($data, $sheetNames, $options);
        $extension = strtolower(substr((string) strrchr($path, '.'), 1));
        $writer = IOFactory::createWriter($spreadsheet, ucfirst($extension));

//        $headers = [
//            'Content-Disposition' => 'attachment; filename='.$filename,
//            'Cache-Control' => 'max-age=0',
//            'Content-Type' => $this->getContentTypeByFilename($filename).'; charset=utf-8',
//        ];

        $writer->save($path);
    }

    public function supports($data): bool
    {
        if (!is_string($data)) {
            return false;
        }

        return file_exists($data);
    }
}
