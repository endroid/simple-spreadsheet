<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet;

use Endroid\SimpleSpreadsheet\Adapter\AdapterInterface;
use Endroid\SimpleSpreadsheet\Adapter\ArrayAdapter;
use Endroid\SimpleSpreadsheet\Adapter\FileAdapter;
use Endroid\SimpleSpreadsheet\Adapter\JsonAdapter;
use Endroid\SimpleSpreadsheet\Adapter\ResponseAdapter;
use Endroid\SimpleSpreadsheet\Adapter\SpreadsheetAdapter;
use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;

final class SimpleSpreadsheet
{
    /** @var array<array<mixed>> */
    private array $sheets = [];

    /** @var array<AdapterInterface<mixed, mixed>> */
    private array $adapters = [];

    public function __construct()
    {
        $this->registerDefaultAdapters();
    }

    private function registerDefaultAdapters(): void
    {
        $this->registerAdapter(new ArrayAdapter());
        $this->registerAdapter(new FileAdapter());
        $this->registerAdapter(new JsonAdapter());
        $this->registerAdapter(new ResponseAdapter());
        $this->registerAdapter(new SpreadsheetAdapter());
    }

    /** @param AdapterInterface<mixed, mixed> $adapter */
    public function registerAdapter(AdapterInterface $adapter): void
    {
        $this->adapters[get_class($adapter)] = $adapter;
    }

    /**
     * @param array<string>|null $sheetNames
     */
    public function load(mixed $data, ?array $sheetNames = null): void
    {
        foreach ($this->adapters as $adapter) {
            if ($adapter->supports($data)) {
                $sheets = $adapter->load($data, $sheetNames);
                $this->append($sheets);
                break;
            }
        }
    }

    /** @param array<string, array<mixed>> $sheets */
    private function append(array $sheets): void
    {
        foreach ($sheets as $sheetName => &$sheetData) {
            if (!isset($this->sheets[$sheetName])) {
                $this->sheets[$sheetName] = [];
            }
            while (count($sheetData) > 0) {
                $sheetDataRow = array_shift($sheetData);
                $this->sheets[$sheetName][] = $sheetDataRow;
            }
        }
    }

    public function createSheet(string $sheetName): void
    {
        if (isset($this->sheets[$sheetName])) {
            throw new SimpleSpreadsheetException(sprintf('Sheet with name "%s" already exists', $sheetName));
        }

        $this->sheets[$sheetName] = [];
    }

    public function renameSheet(string $sourceName, string $targetName): void
    {
        $this->duplicateSheet($sourceName, $targetName);
        unset($this->sheets[$sourceName]);
    }

    public function duplicateSheet(string $sourceName, string $targetName): void
    {
        if (!isset($this->sheets[$sourceName])) {
            throw new SimpleSpreadsheetException(sprintf('Sheet with name "%s" does not exist', $sourceName));
        }

        $this->sheets[$targetName] = $this->sheets[$sourceName];
    }

    public function removeSheet(string $sheetName): void
    {
        if (!isset($this->sheets[$sheetName])) {
            throw new SimpleSpreadsheetException(sprintf('Sheet with name "%s" does not exist', $sheetName));
        }

        unset($this->sheets[$sheetName]);
    }

    /**
     * @param array<string>|null $sheetNames
     * @param array<mixed>       $options
     */
    public function save(string $adapterClass, ?array $sheetNames = null, array $options = []): mixed
    {
        $adapter = $this->adapters[$adapterClass];

        if (!$adapter instanceof AdapterInterface) {
            throw new SimpleSpreadsheetException(sprintf('Adapter class "%s" not found', $adapterClass));
        }

        return $adapter->save($this->sheets, $sheetNames, $options);
    }
}
