<?php

declare(strict_types=1);

namespace Endroid\SimpleSpreadsheet\Adapter;

use Endroid\SimpleSpreadsheet\Exception\SimpleSpreadsheetException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

/**
 * @implements AdapterInterface<Response, Response>
 */
final readonly class ResponseAdapter implements AdapterInterface
{
    public function __construct(
        private SpreadsheetAdapter $spreadsheetAdapter = new SpreadsheetAdapter(),
    ) {}

    /** @var array<string, string> */
    private const array CONTENT_TYPES_BY_EXTENSION = [
        'csv' => 'text/csv',
        'xls' => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    #[\Override]
    public function supports(mixed $data): bool
    {
        throw new SimpleSpreadsheetException('Unable to load from response');
    }

    #[\Override]
    public function load(mixed $data, ?array $sheetNames = null): array
    {
        throw new SimpleSpreadsheetException('Unable to load from response');
    }

    #[\Override]
    public function save(array $data, ?array $sheetNames = null, array $options = []): Response
    {
        if (!is_string($options['filename'] ?? null)) {
            throw new SimpleSpreadsheetException('Please specify the filename via options');
        }

        $filename = $options['filename'];
        $spreadsheet = $this->spreadsheetAdapter->save($data, $sheetNames, $options);
        $extension = strtolower(substr((string) strrchr($filename, needle: '.'), offset: 1));
        if (!array_key_exists($extension, self::CONTENT_TYPES_BY_EXTENSION)) {
            throw new SimpleSpreadsheetException(sprintf('Unsupported response extension "%s"', $extension));
        }
        $writer = IOFactory::createWriter($spreadsheet, ucfirst($extension));

        ob_start();
        $writer->save('php://output');
        $contents = (string) ob_get_clean();

        $response = new Response($contents);
        $response->headers->add([
            'Content-Type' => self::CONTENT_TYPES_BY_EXTENSION[$extension],
            'Content-Disposition' => $response->headers->makeDisposition(
                ResponseHeaderBag::DISPOSITION_INLINE,
                $filename,
            ),
        ]);

        return $response;
    }
}
