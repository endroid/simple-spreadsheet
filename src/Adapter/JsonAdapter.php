<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;

final class JsonAdapter extends AbstractAdapter
{
    private ArrayAdapter $arrayAdapter;

    public function __construct()
    {
        $this->arrayAdapter = new ArrayAdapter();
    }

    public function load($data, array $sheetNames = null): array
    {
        $data = json_decode($data, true);

        return $this->arrayAdapter->load($data, $sheetNames);
    }

    public function save(array $data, array $sheetNames = null, array $options = []): string
    {
        $data = json_encode($data);
        if (!is_string($data)) {
            throw new SimpleSpreadsheetException('Unable to encode data to JSON');
        }

        return $data;
    }

    public function supports($data): bool
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
}
