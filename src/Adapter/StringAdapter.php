<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

class StringAdapter extends AbstractAdapter
{
    public function load($data, array $sheetNames = null): array
    {
    }

    public function save(array $data, array $sheetNames = null)
    {
        ob_start();

        $this->saveToOutput($filename, $sheetNames, false);

        return (string) ob_get_clean();
    }

    public function supports($data): bool
    {
        return is_string($data);
    }
}
