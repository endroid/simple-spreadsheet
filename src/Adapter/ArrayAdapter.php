<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

class ArrayAdapter extends AbstractAdapter
{
    public function load($data, array $sheetNames = null): array
    {
        $sheets = [];

        foreach ($data as $sheetName => $sheetData) {
            if (null === $sheetNames || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheetData;
            }
        }

        return $sheets;
    }

    public function save(array $data, array $sheetNames = null, array $options = [])
    {
        $sheets = [];

        foreach ($data as $sheetName => $sheetData) {
            if (null === $sheetNames || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheetData;
            }
        }

        return $sheets;
    }

    public function supports($data): bool
    {
        if (!is_array($data)) {
            return false;
        }

        foreach ($data as $sheet) {
            if (!is_array($sheet)) {
                return false;
            }
        }

        return true;
    }
}
