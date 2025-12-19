<?php

namespace AlexMarquezini\ReportEngine;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelGenerator
{
    protected ReportDefinition $definition;
    protected array $processedData;
    protected Spreadsheet $spreadsheet;
    protected ?string $templatePath = null;

    // Template row definitions (indices)
    protected ?int $itemRowIndex = null;
    protected ?int $groupHeaderRowIndex = null;
    protected ?int $totalRowIndex = null;

    // Blueprints
    protected array $itemBlueprint = [];
    protected array $groupHeaderBlueprint = [];
    protected array $totalBlueprint = [];

    public function __construct(array $processedData, ReportDefinition $definition)
    {
        $this->processedData = $processedData;
        $this->definition = $definition;
        $this->spreadsheet = new Spreadsheet();
    }

    public function setTemplate(string $path): self
    {
        if (file_exists($path)) {
            $this->templatePath = $path;
            $this->spreadsheet = IOFactory::load($path);
        }
        return $this;
    }

    public function generate(): Spreadsheet
    {
        if ($this->templatePath) {
            return $this->generateFromTemplate();
        }

        return $this->generateDefault();
    }

    protected function generateFromTemplate(): Spreadsheet
    {
        $sheet = $this->spreadsheet->getActiveSheet();

        // 1. Analyze Template to find blueprint rows
        $this->analyzeTemplate($sheet);

        // 2. Global Replacements (Title, etc)
        $this->replaceGlobalTags($sheet);

        // 3. Extract Blueprints
        if ($this->itemRowIndex) {
            $this->itemBlueprint = $this->extractRowBlueprint($sheet, $this->itemRowIndex);
        }
        if ($this->groupHeaderRowIndex) {
            $this->groupHeaderBlueprint = $this->extractRowBlueprint($sheet, $this->groupHeaderRowIndex);
        }
        if ($this->totalRowIndex) {
            $this->totalBlueprint = $this->extractRowBlueprint($sheet, $this->totalRowIndex);
        }

        // 4. Remove blueprint rows from sheet (reverse order to keep indices valid during deletion)
        $rowsToDelete = array_filter([$this->itemRowIndex, $this->groupHeaderRowIndex, $this->totalRowIndex]);
        rsort($rowsToDelete);
        foreach ($rowsToDelete as $r) {
            $sheet->removeRow($r);
        }

        // 5. Determine Start Row
        $startRow = $rowsToDelete ? min($rowsToDelete) : 1;
        $currentRow = $startRow;

        $this->renderGroupRecursiveTemplate($sheet, $this->processedData['data'], $currentRow);

        return $this->spreadsheet;
    }

    protected function extractRowBlueprint(Worksheet $sheet, int $rowIndex): array
    {
        $blueprint = [];
        $highestCol = $sheet->getHighestColumn();
        $highestColIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestCol);

        for ($col = 1; $col <= $highestColIndex; $col++) {
            $colString = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col);
            $cell = $sheet->getCell($colString . $rowIndex);
            $blueprint[$colString] = [
                'value' => $cell->getValue(),
                'style' => $sheet->getStyle($colString . $rowIndex)->exportArray()
            ];
        }

        // Extract Merges
        $blueprint['_merges'] = [];
        foreach ($sheet->getMergeCells() as $merge) {
            $range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($merge);
            $start = $range[0][0];
            $end = $range[0][1];

            $startCoord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($start);
            $endCoord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($end);

            // Check if this merge is strictly on the current row
            if ($startCoord[1] == $rowIndex && $endCoord[1] == $rowIndex) {
                $blueprint['_merges'][] = $startCoord[0] . ':' . $endCoord[0];
            }
        }

        return $blueprint;
    }

    protected function renderGroupRecursiveTemplate(Worksheet $sheet, array $items, int &$row)
    {
        foreach ($items as $item) {
            if (is_array($item) && isset($item['group_field'])) {
                // --- Group Header ---
                $sheet->insertNewRowBefore($row, 1);

                if (!empty($this->groupHeaderBlueprint)) {
                    $this->applyBlueprint($sheet, $row, $this->groupHeaderBlueprint, $item, 'group');
                } else {
                    // Fallback default style
                    $sheet->setCellValue("A$row", "Agrupamento: " . $item['group_value']);
                    $sheet->getStyle("A$row")->getFont()->setBold(true);
                }
                $row++;

                // --- Children ---
                $this->renderGroupRecursiveTemplate($sheet, $item['items'], $row);

                // --- Group Total ---
                $sheet->insertNewRowBefore($row, 1);

                if (!empty($this->totalBlueprint)) {
                    $this->applyBlueprint($sheet, $row, $this->totalBlueprint, $item, 'total');
                } else {
                    // Fallback default style
                    $sheet->setCellValue("A$row", "Total " . $item['group_value']);
                }
                $row++;

            } else {
                // --- Item Row ---
                $sheet->insertNewRowBefore($row, 1);

                if (!empty($this->itemBlueprint)) {
                    $this->applyBlueprint($sheet, $row, $this->itemBlueprint, $item, 'item');
                } else {
                    // Fallback default style (simplified)
                    $col = 'A';
                    foreach ($this->definition->getColumns() as $field => $config) {
                        $val = is_object($item) ? ($item->$field ?? '') : ($item[$field] ?? '');
                        $sheet->setCellValue("$col$row", $val);
                        $col++;
                    }
                }
                $row++;
            }
        }
    }

    protected function applyBlueprint(Worksheet $sheet, int $row, array $blueprint, $dataItem, string $type)
    {
        foreach ($blueprint as $col => $cellData) {
            if ($col === '_merges')
                continue;

            $sheet->getCell($col . $row)->setValue($cellData['value']);
            $sheet->getStyle($col . $row)->applyFromArray($cellData['style']);

            // Interpolation
            $val = $sheet->getCell($col . $row)->getValue();
            if (is_string($val) && strpos($val, '{{') !== false) {

                // 1. Group Tags
                if ($type === 'group') {
                    $val = str_replace('{{group_value}}', $dataItem['group_value'], $val);
                    $val = str_replace('{{group_field}}', $dataItem['group_field'], $val);
                }

                // 2. Total Tags
                if ($type === 'total') {
                    foreach ($dataItem['totals'] as $field => $amount) {
                        $tag = '{{total_' . $field . '}}';
                        if (strpos($val, $tag) !== false) {
                            $formatted = number_format((float) $amount, 2, ',', '.');
                            $val = str_replace($tag, $formatted, $val);
                        }
                    }
                    $val = str_replace('{{group_value}}', $dataItem['group_value'], $val);
                }

                // 3. Item Tags
                if ($type === 'item') {
                    foreach ($this->definition->getColumns() as $field => $config) {
                        $tag = '{{' . $field . '}}';
                        if (strpos($val, $tag) !== false) {
                            $dataVal = is_object($dataItem) ? ($dataItem->$field ?? '') : ($dataItem[$field] ?? '');

                            // Format
                            if (isset($config['format']) && $config['format'] === 'currency') {
                                $dataVal = number_format((float) $dataVal, 2, ',', '.');
                            } elseif (isset($config['format']) && $config['format'] === 'date') {
                                $dataVal = date('d/m/Y', strtotime($dataVal));
                            }
                            $val = str_replace($tag, $dataVal, $val);

                            // Handle Actions (Hyperlinks)
                            if (isset($config['action']) && isset($config['action']['route'])) {
                                $url = $config['action']['route'];
                                if (preg_match_all('/\{(\w+)\}/', $url, $matches)) {
                                    foreach ($matches[1] as $param) {
                                        $paramVal = is_object($dataItem) ? ($dataItem->$param ?? '') : ($dataItem[$param] ?? '');
                                        $url = str_replace('{' . $param . '}', $paramVal, $url);
                                    }
                                }
                                $sheet->getCell($col . $row)->getHyperlink()->setUrl($url);
                                $sheet->getCell($col . $row)->getHyperlink()->setTooltip('Clique para ver detalhes');
                                $sheet->getStyle($col . $row)->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLUE));
                                $sheet->getStyle($col . $row)->getFont()->setUnderline(true);
                            }
                        }
                    }
                }

                $sheet->getCell($col . $row)->setValue($val);
            }
        }

        // Apply Merges
        if (isset($blueprint['_merges'])) {
            foreach ($blueprint['_merges'] as $mergeCols) {
                $parts = explode(':', $mergeCols);
                $sheet->mergeCells($parts[0] . $row . ':' . $parts[1] . $row);
            }
        }
    }

    protected function analyzeTemplate(Worksheet $sheet)
    {
        $highestRow = $sheet->getHighestRow();
        $columns = array_keys($this->definition->getColumns());

        for ($row = 1; $row <= $highestRow; $row++) {
            $rowIterator = $sheet->getRowIterator($row)->current();
            if (!$rowIterator)
                continue;

            $cellIterator = $rowIterator->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);

            $isItemRow = false;
            $isGroupRow = false;
            $isTotalRow = false;

            foreach ($cellIterator as $cell) {
                $val = $cell->getValue();
                if (is_string($val)) {
                    foreach ($columns as $col) {
                        if (strpos($val, '{{' . $col . '}}') !== false) {
                            $isItemRow = true;
                        }
                    }
                    if (strpos($val, '{{total_') !== false) {
                        $isTotalRow = true;
                    } elseif (strpos($val, '{{group_value}}') !== false) {
                        $isGroupRow = true;
                    }
                }
            }

            if ($isItemRow && !$this->itemRowIndex)
                $this->itemRowIndex = $row;
            if ($isGroupRow && !$this->groupHeaderRowIndex)
                $this->groupHeaderRowIndex = $row;
            if ($isTotalRow && !$this->totalRowIndex)
                $this->totalRowIndex = $row;
        }
    }

    protected function replaceGlobalTags(Worksheet $sheet)
    {
        $limit = min(20, $sheet->getHighestRow());

        for ($row = 1; $row <= $limit; $row++) {
            $rowIterator = $sheet->getRowIterator($row)->current();
            if (!$rowIterator)
                continue;

            $cellIterator = $rowIterator->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);

            foreach ($cellIterator as $cell) {
                $val = $cell->getValue();
                if (is_string($val)) {
                    $val = str_replace('{{titulo}}', $this->definition->getTitle(), $val);
                    foreach ($this->definition->getParameters() as $key => $value) {
                        $val = str_replace('{{' . $key . '}}', $value, $val);
                    }
                    $cell->setValue($val);
                }
            }
        }
    }

    protected function generateDefault(): Spreadsheet
    {
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->setTitle(substr($this->definition->getTitle(), 0, 31));

        $row = 1;
        $sheet->setCellValue("A$row", $this->definition->getTitle());
        $sheet->mergeCells("A$row:F$row");
        $sheet->getStyle("A$row")->getFont()->setBold(true)->setSize(14);
        $row += 2;

        $col = 'A';
        foreach ($this->definition->getColumns() as $column) {
            $sheet->setCellValue("$col$row", $column['label']);
            $sheet->getStyle("$col$row")->getFont()->setBold(true);
            if (isset($column['width'])) {
                $sheet->getColumnDimension($col)->setWidth($column['width']);
            } else {
                $sheet->getColumnDimension($col)->setAutoSize(true);
            }
            $col++;
        }
        $row++;

        $this->renderGroup($sheet, $this->processedData['data'], $row);
        $this->renderTotals($sheet, $this->processedData['grand_totals'], $row, "Total Geral");

        return $this->spreadsheet;
    }

    private function renderGroup($sheet, array $items, int &$row)
    {
        foreach ($items as $item) {
            if (is_array($item) && isset($item['group_field'])) {
                $sheet->setCellValue("A$row", "Agrupamento: " . $item['group_value']);
                $sheet->getStyle("A$row")->getFont()->setBold(true);
                $sheet->getStyle("A$row")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFEEEEEE');
                $row++;

                $this->renderGroup($sheet, $item['items'], $row);
                $this->renderTotals($sheet, $item['totals'], $row, "Total " . $item['group_value']);
                $row++;
            } else {
                $col = 'A';
                foreach ($this->definition->getColumns() as $field => $config) {
                    $val = is_object($item) ? ($item->$field ?? '') : ($item[$field] ?? '');
                    if (isset($config['format'])) {
                        if ($config['format'] === 'currency') {
                            $val = number_format((float) $val, 2, ',', '.');
                            $sheet->getStyle("$col$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                        } elseif ($config['format'] === 'date') {
                            $val = date('d/m/Y', strtotime($val));
                        }
                    }
                    $sheet->setCellValue("$col$row", $val);
                    $col++;
                }
                $row++;
            }
        }
    }

    private function renderTotals($sheet, array $totals, int &$row, string $label)
    {
        $sheet->setCellValue("A$row", $label);
        $sheet->getStyle("A$row")->getFont()->setBold(true);
        $sheet->getStyle("A$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        $col = 'A';
        foreach ($this->definition->getColumns() as $field => $config) {
            if (in_array($field, array_keys($totals))) {
                $val = number_format($totals[$field], 2, ',', '.');
                $sheet->setCellValue("$col$row", $val);
                $sheet->getStyle("$col$row")->getFont()->setBold(true);
                $sheet->getStyle("$col$row")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            }
            $col++;
        }
        $row++;
    }
}
