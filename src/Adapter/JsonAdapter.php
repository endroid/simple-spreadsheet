<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet\Adapter;

class JsonAdapter extends ArrayAdapter
{
    public function load($data, array $sheetNames = null): array
    {
        $array = json_decode($data, true);

        return parent::load($data, $sheetNames);
    }

    public function save(array $data, array $sheetNames = null, array $options = [])
    {
        return json_encode($data);
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
