<?php

namespace AlexMarquezini\ReportEngine;

use Knp\Snappy\Pdf;

class PdfGenerator
{
    protected ReportDefinition $definition;
    protected array $processedData;
    protected Pdf $snappy;

    public function __construct(array $processedData, ReportDefinition $definition)
    {
        $this->processedData = $processedData;
        $this->definition = $definition;

        // Note: ROOTPATH and WRITEPATH are CodeIgniter constants.
        // For a framework-agnostic package, these should be passed via configuration.
        // Keeping them for now as per "don't touch what works" instruction for this specific project context.
        $this->snappy = new Pdf(
            ROOTPATH . 'vendor\h4cc\wkhtmltopdf-amd64\bin\wkhtmltopdf',
            [
                'margin-top' => 5,
                'margin-right' => 5,
                'margin-bottom' => 10,
                'margin-left' => 5,
                'orientation' => 'Landscape',
                'page-size' => 'A4',
                'outline' => false,
                'footer-line' => true,
                'footer-spacing' => 2,
                'footer-font-size' => 8,
                'footer-right' => 'Página [page] de [toPage]',
                'footer-font-name' => 'sans-serif',
            ]
        );

        $this->snappy->setTemporaryFolder(WRITEPATH . 'uploads/temp');
        $this->snappy->setTimeout(300);
    }

    public function generate(): string
    {
        // Note: view() is a CodeIgniter helper. 
        // Ideally this should use a renderer interface.
        $html = view('reports/templates/default_pdf', [
            'definition' => $this->definition,
            'data' => $this->processedData['data'],
            'grand_totals' => $this->processedData['grand_totals'],
            'metadata' => [
                'title' => $this->definition->getTitle(),
                'parameters' => $this->definition->getParameters(),
                'columns' => $this->definition->getColumns(),
            ]
        ]);

        return $this->snappy->getOutputFromHtml($html);
    }
}
