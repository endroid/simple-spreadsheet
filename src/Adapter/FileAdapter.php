<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * @implements AdapterInterface<string, null>
 */
final readonly class FileAdapter implements AdapterInterface
{
    public function __construct(
        private SpreadsheetAdapter $spreadsheetAdapter = new SpreadsheetAdapter(),
    ) {
    }

    public function supports(mixed $data): bool
    {
        if (!is_string($data)) {
            return false;
        }

        return file_exists($data);
    }

    public function load(mixed $data, ?array $sheetNames = null): array
    {
        $spreadsheet = IOFactory::load($data);

        return $this->spreadsheetAdapter->load($spreadsheet, $sheetNames);
    }

    public function save(array $data, ?array $sheetNames = null, array $options = []): mixed
    {
        if (!isset($options['path'])) {
            throw new SimpleSpreadsheetException('Please specify the output path via options');
        }

        $path = $options['path'];
        $spreadsheet = $this->spreadsheetAdapter->save($data, $sheetNames, $options);
        $extension = strtolower(substr((string) strrchr($path, '.'), 1));
        $writer = IOFactory::createWriter($spreadsheet, ucfirst($extension));

        $writer->save($path);

        return null;
    }
}
