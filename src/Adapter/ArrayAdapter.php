<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

final class ArrayAdapter extends AbstractAdapter
{
    /** @param array<string, array<mixed>> $data */
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

    /** @return array<mixed> */
    public function save(array $data, array $sheetNames = null, array $options = []): array
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
