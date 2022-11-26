<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\IOFactory;

class FileAdapter extends SpreadsheetAdapter
{
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

        $writer->save($path);

        return null;
    }

    public function supports($data): bool
    {
        if (!is_string($data)) {
            return false;
        }

        return file_exists($data);
    }
}
