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
    ) {}

    #[\Override]
    public function supports(mixed $data): bool
    {
        if (!is_string($data)) {
            return false;
        }

        return $this->arrayAdapter->supports(json_decode($data, associative: true));
    }

    #[\Override]
    public function load(mixed $data, ?array $sheetNames = null): array
    {
        return $this->arrayAdapter->load(json_decode($data, associative: true), $sheetNames);
    }

    #[\Override]
    public function save(array $data, ?array $sheetNames = null, array $options = []): mixed
    {
        $data = json_encode($data);
        if (!is_string($data)) {
            throw new SimpleSpreadsheetException('Unable to encode data to JSON');
        }

        return $data;
    }
}
