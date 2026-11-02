<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

/**
 * @implements AdapterInterface<array<array-key, array<mixed>>, array<array-key, array<mixed>>>
 */
final readonly class ArrayAdapter implements AdapterInterface
{
    /** @phpstan-assert-if-true array<array-key, array<mixed>> $data */
    #[\Override]
    public function supports(mixed $data): bool
    {
        if (!is_array($data)) {
            return false;
        }

        foreach (array_keys($data) as $key) {
            if (!is_array($data[$key] ?? null)) {
                return false;
            }
        }

        return true;
    }

    /** @param mixed $data */
    #[\Override]
    public function load(mixed $data, ?array $sheetNames = null): array
    {
        if (!$this->supports($data)) {
            throw new \InvalidArgumentException('Expected an array of sheets');
        }
        $sheets = [];

        foreach ($data as $sheetName => $sheetData) {
            if (!(null === $sheetNames || in_array((string) $sheetName, $sheetNames, strict: true))) {
                continue;
            }

            $sheets[$sheetName] = $sheetData;
        }

        return $sheets;
    }

    /** @return array<array-key, array<mixed>> */
    #[\Override]
    public function save(array $data, ?array $sheetNames = null, array $options = []): array
    {
        $sheets = [];

        foreach ($data as $sheetName => $sheetData) {
            if (!(null === $sheetNames || in_array((string) $sheetName, $sheetNames, strict: true))) {
                continue;
            }

            $sheets[$sheetName] = $sheetData;
        }

        return $sheets;
    }
}
