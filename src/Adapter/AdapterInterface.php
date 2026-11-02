<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

/**
 * @template T
 * @template U
 */
interface AdapterInterface
{
    public function supports(mixed $data): bool;

    /**
     * @param T                  $data
     * @param array<string>|null $sheetNames
     *
     * @return array<array-key, array<mixed>>
     */
    public function load(mixed $data, ?array $sheetNames = null): array;

    /**
     * @param array<array-key, array<mixed>> $data
     * @param array<string>|null          $sheetNames
     * @param array<mixed>                $options
     *
     * @return U
     */
    public function save(array $data, ?array $sheetNames = null, array $options = []): mixed;
}
