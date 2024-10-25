<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

/**
 * @implements AdapterInterface<array<string, array<mixed>>, array<mixed>>
 */
final readonly class ArrayAdapter implements AdapterInterface
{
    public function supports(mixed $data): bool
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

    public function load(mixed $data, ?array $sheetNames = null): array
    {
        $sheets = [];

        foreach ($data as $sheetName => $sheetData) {
            if (null === $sheetNames || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheetData;
            }
        }

        return $sheets;
    }

    public function save(array $data, ?array $sheetNames = null, array $options = []): array
    {
        $sheets = [];

        foreach ($data as $sheetName => $sheetData) {
            if (null === $sheetNames || in_array($sheetName, $sheetNames)) {
                $sheets[$sheetName] = $sheetData;
            }
        }

        return $sheets;
    }
}
