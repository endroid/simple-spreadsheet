<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

interface AdapterInterface
{
    /**
     * @param mixed              $data
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
