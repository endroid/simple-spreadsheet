<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;

/**
 * @implements AdapterInterface<string, string>
 */
final readonly class JsonAdapter implements AdapterInterface
{
    public function __construct(
        private ArrayAdapter $arrayAdapter = new ArrayAdapter(),
    ) {
    }

    public function supports(mixed $data): bool
    {
        if (!is_string($data)) {
            return false;
        }

        $data = json_decode($data, true);

        if (!is_array($data)) {
            return false;
        }

        return true;
    }

    public function load(mixed $data, ?array $sheetNames = null): array
    {
        $data = json_decode($data, true);

        return $this->arrayAdapter->load($data, $sheetNames);
    }

    public function save(array $data, ?array $sheetNames = null, array $options = []): mixed
    {
        $data = json_encode($data);
        if (!is_string($data)) {
            throw new SimpleSpreadsheetException('Unable to encode data to JSON');
        }

        return $data;
    }
}
