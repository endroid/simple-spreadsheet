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
    public function load($data, array $filterSheets = null): array;

    public function save(array $data, array $filterSheets = null);

    public function supports($data): bool;

    public function getPriority(): int;
}
