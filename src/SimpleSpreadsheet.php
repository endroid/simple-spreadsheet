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
use Endroid\SimpleSpreadsheet\Adapter\SpreadsheetAdapter;
use Endroid\SimpleSpreadsheet\Adapter\StringAdapter;
use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

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
                $adapter->load($data, $sheetNames);
                break;
            }
        }
    }

    public function save(string $adapterClass, array $sheetNames = null)
    {
        $adapter = $this->adapters[$adapterClass];

        if (!$adapter instanceof AdapterInterface) {
            throw new SimpleSpreadsheetException(sprintf('Adapter class "%s" not found', $adapterClass));
        }

        return $adapter->save($this->sheets, $sheetNames);
    }

    public function loadFromArray(array $data): void
    {
        // If the data is not multidimensional make it so
    }

    public function loadFromSpreadsheet(Spreadsheet $excel, array $sheetNames = []): void
    {
        foreach ($excel->getWorksheetIterator() as $sheet) {
            if (0 == count($sheetNames) || in_array($sheet->getTitle(), $sheetNames)) {
                $this->loadFromSheet($sheet);
            }
        }
    }

    public function loadFromSheet(Worksheet $excelSheet): void
    {
        $sheet = [];

        $sheetData = $excelSheet->toArray('', false, false);

        // Remove possible empty leading rows
        while ('' == $sheetData[0][0] && count($sheetData) > 0) {
            array_shift($sheetData);
        }

        // First row always contains the headers
        $columns = array_shift($sheetData);

        // Remove headers from the end until first name is found
        for ($i = count($columns) - 1; $i >= 0; --$i) {
            if ('' == $columns[$i]) {
                unset($columns[$i]);
            } else {
                break;
            }
        }

        // Next rows contain the actual data
        foreach ($sheetData as $row) {
            // Ignore empty rows
            if ('' == trim(implode('', $row))) {
                continue;
            }

            // Map data to column names
            $associativeRow = [];
            foreach ($row as $key => $value) {
                if (!isset($columns[$key])) {
                    continue;
                }
                if ('null' == strtolower($value)) {
                    $value = null;
                }
                $associativeRow[$columns[$key]] = $value;
            }
            $sheet[] = $associativeRow;
        }

        $this->sheets[$excelSheet->getTitle()] = $sheet;
    }

    public function saveToArray(array $sheetNames = []): array
    {
    }

    public function saveToOutput(string $filename, array $sheetNames = [], bool $setHeaders = true): void
    {
        if ($setHeaders) {
            $headers = $this->getHeadersByFilename($filename);
            foreach ($headers as $key => $value) {
                header($key.': '.$value);
            }
        }

        $writer = $this->getWriterByFilename($filename, $sheetNames);
        $writer->save('php://output');
    }

    public function saveToResponse(string $filename, array $sheetNames = []): Response
    {
        $responseClass = 'Symfony\Component\HttpFoundation\Response';
        if (!class_exists($responseClass)) {
            throw new SimpleSpreadsheetException('Class "'.$responseClass.'" not found: make sure symfony/http-foundation is installed');
        }

        $response = new Response($this->saveToString($filename, $sheetNames));
        $response->headers->add([
            'Content-Type' => $this->getHeadersByFilename($filename),
            'Content-Disposition' => $response->headers->makeDisposition(ResponseHeaderBag::DISPOSITION_INLINE, $filename),
        ]);

        return $response;
    }

    protected function getContentTypeByFilename($filename): string
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->contentTypes[$extension])) {
            throw new Exception(sprintf('No content type defined for file extension "%s"', $extension));
        }

        return $this->contentTypes[$extension];
    }

    public function getWriterByFilename(string $filename, array $sheetNames = []): IWriter
    {
        $extension = $this->getExtension($filename);

        if (!isset($this->writers[$extension])) {
            throw new Exception(sprintf('No writer defined for file extension "%s"', $extension));
        }

        $excel = $this->saveToExcel($sheetNames);

        return IOFactory::createWriter($excel, $this->writers[$extension]);
    }

    public function getHeadersByFilename(string $filename): array
    {
        $headers = [
            'Content-Disposition' => 'attachment; filename='.$filename,
            'Cache-Control' => 'max-age=0',
            'Content-Type' => $this->getContentTypeByFilename($filename).'; charset=utf-8',
        ];

        return $headers;
    }

    private function getExtension(string $filename): string
    {
        return strtolower(substr((string) strrchr($filename, '.'), 1));
    }
}
