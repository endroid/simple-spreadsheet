<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class ResponseAdapter extends SpreadsheetAdapter
{
    /** @var array<string, string> */
    private array $contentTypesByExtension = [
        'csv' => 'text/csv',
        'xls' => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    public function load($data, array $sheetNames = null): array
    {
        throw new SimpleSpreadsheetException('Unable to load from response');
    }

    public function save(array $data, array $sheetNames = null, array $options = [])
    {
        if (!isset($options['filename'])) {
            throw new SimpleSpreadsheetException('Please specify the filename via options');
        }

        $filename = $options['filename'];
        $spreadsheet = parent::save($data, $sheetNames, $options);
        $extension = strtolower(substr((string) strrchr($filename, '.'), 1));
        $writer = IOFactory::createWriter($spreadsheet, ucfirst($extension));

        ob_start();
        $writer->save('php://output');
        $contents = (string) ob_get_clean();

        $response = new Response($contents);
        $response->headers->add([
            'Content-Type' => $this->contentTypesByExtension[$extension],
            'Content-Disposition' => $response->headers->makeDisposition(ResponseHeaderBag::DISPOSITION_INLINE, $filename),
        ]);

        return $response;
    }

    public function supports($data): bool
    {
        throw new SimpleSpreadsheetException('Unable to load from response');
    }
}
