<?php

declare(strict_types=1);

/*
 * (c) Jeroen van den Enden <info@endroid.nl>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace Endroid\SimpleSpreadsheet;

use Endroid\SimpleSpreadsheet\Adapter\AdapterInterface;
use Endroid\SimpleSpreadsheet\Adapter\ArrayAdapter;
use Endroid\SimpleSpreadsheet\Adapter\FileAdapter;
use Endroid\SimpleSpreadsheet\Adapter\ResponseAdapter;
use Endroid\SimpleSpreadsheet\Adapter\SpreadsheetAdapter;
use Endroid\SimpleSpreadsheet\Adapter\StringAdapter;
use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;

class SimpleSpreadsheet
{
    /** @var array[] */
    private $sheets = [];

    /** @var AdapterInterface[] */
    private $adapters = [];

    public function __construct()
    {
        $this->registerDefaultAdapters();
    }

    private function registerDefaultAdapters(): void
    {
        $this->registerAdapter(new ArrayAdapter());
        $this->registerAdapter(new FileAdapter());
        $this->registerAdapter(new ResponseAdapter());
        $this->registerAdapter(new SpreadsheetAdapter());
        $this->registerAdapter(new StringAdapter());
    }

    public function registerAdapter(AdapterInterface $adapter): void
    {
        $this->adapters[get_class($adapter)] = $adapter;
        $this->prioritizeAdapters();
    }

    private function prioritizeAdapters(): void
    {
        uasort($this->adapters, function (AdapterInterface $adapter1, AdapterInterface $adapter2) {
            return $adapter2->getPriority() - $adapter1->getPriority();
        });
    }

    public function load($data, array $sheetNames = null): void
    {
        foreach ($this->adapters as $adapter) {
            if ($adapter->supports($data)) {
                $sheets = $adapter->load($data, $sheetNames);
                $this->append($sheets);
                break;
            }
        }
    }

    private function append(array $sheets): void
    {
        foreach ($sheets as $sheetName => &$sheetData) {
            while (count($sheetData) > 0) {
                $sheetDataRow = array_shift($sheetData);
                $this->sheets[$sheetName][] = $sheetDataRow;
            }
        }
    }

    public function save(string $adapterClass, array $sheetNames = null, array $options = [])
    {
        $adapter = $this->adapters[$adapterClass];

        if (!$adapter instanceof AdapterInterface) {
            throw new SimpleSpreadsheetException(sprintf('Adapter class "%s" not found', $adapterClass));
        }

        return $adapter->save($this->sheets, $sheetNames, $options);
    }
}
