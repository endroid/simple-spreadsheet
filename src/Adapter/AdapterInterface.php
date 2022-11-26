<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

interface AdapterInterface
{
    /**
     * @template T
     *
     * @param T                  $data
     * @param array<string>|null $sheetNames
     *
     * @return array<string, array<mixed>>
     */
    public function load($data, array $sheetNames = null): array;

    /**
     * @param array<string, array<mixed>> $data
     * @param array<string>|null          $sheetNames
     * @param array<mixed>                $options
     *
     * @return mixed
     */
    public function save(array $data, array $sheetNames = null, array $options = []);

    /** @param mixed $data */
    public function supports($data): bool;

    public function getPriority(): int;
}
