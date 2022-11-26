<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

abstract class AbstractAdapter implements AdapterInterface
{
    public function getPriority(): int
    {
        return 1;
    }
}
