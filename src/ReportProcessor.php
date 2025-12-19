<?php

namespace AlexMarquezini\ReportEngine;

class ReportProcessor
{
    protected ReportDefinition $definition;
    protected array $data;

    public function __construct(array $data, ReportDefinition $definition)
    {
        $this->data = $data;
        $this->definition = $definition;
    }

    /**
     * Processes the flat data into a structured format with totals.
     */
    public function process(): array
    {
        $groups = $this->definition->getGroupBy();
        $totalizers = $this->definition->getTotalizers();

        $structuredData = $this->groupData($this->data, $groups, $totalizers);

        // Calculate Grand Totals
        $grandTotals = $this->calculateTotals($this->data, $totalizers);

        return [
            'metadata' => [
                'title' => $this->definition->getTitle(),
                'parameters' => $this->definition->getParameters(),
                'columns' => $this->definition->getColumns(),
            ],
            'data' => $structuredData,
            'grand_totals' => $grandTotals
        ];
    }

    private function groupData(array $rows, array $groupFields, array $totalizers): array
    {
        if (empty($groupFields)) {
            return $rows; // No more grouping, return raw rows
        }

        // Get the first key/value pair
        $fieldKey = array_key_first($groupFields);
        $fieldValue = $groupFields[$fieldKey];

        // Remove it from the array for recursion
        unset($groupFields[$fieldKey]);

        $groupByField = is_int($fieldKey) ? $fieldValue : $fieldKey;
        $displayField = is_int($fieldKey) ? $groupByField : $fieldValue;

        $grouped = [];

        foreach ($rows as $row) {
            // Handle object or array access
            $key = is_object($row) ? ($row->$groupByField ?? '') : ($row[$groupByField] ?? '');
            $displayVal = is_object($row) ? ($row->$displayField ?? '') : ($row[$displayField] ?? '');

            if (!isset($grouped[$key])) {
                $grouped[$key] = [
                    'group_field' => $groupByField,
                    'group_value' => $displayVal, // Use display value for label
                    'group_key' => $key,          // Keep original code
                    'items' => [],
                    'raw_rows' => [] // Temp storage for recursion
                ];
            }
            $grouped[$key]['raw_rows'][] = $row;
        }

        // Recursively process subgroups and calculate totals for this level
        foreach ($grouped as $key => &$group) {
            $group['totals'] = $this->calculateTotals($group['raw_rows'], $totalizers);

            // Recurse
            $group['items'] = $this->groupData($group['raw_rows'], $groupFields, $totalizers);

            // Cleanup temp storage
            unset($group['raw_rows']);
        }

        return array_values($grouped);
    }

    private function calculateTotals(array $rows, array $fields): array
    {
        $totals = array_fill_keys($fields, 0.0);

        foreach ($rows as $row) {
            foreach ($fields as $field) {
                $val = is_object($row) ? ($row->$field ?? 0) : ($row[$field] ?? 0);
                $totals[$field] += (float) $val;
            }
        }

        return $totals;
    }
}
