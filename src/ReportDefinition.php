<?php

namespace AlexMarquezini\ReportEngine;

class ReportDefinition
{
    /**
     * @var array List of columns to display.
     * Format: ['field_name' => ['label' => 'Label', 'format' => 'currency|date|string', 'width' => 15]]
     */
    protected array $columns = [];

    /**
     * @var array List of fields to group by.
     * Order matters: first field is the top-level group.
     */
    protected array $groupBy = [];

    /**
     * @var array List of fields to calculate totals for.
     */
    protected array $totalizers = [];

    /**
     * @var string Title of the report.
     */
    protected string $title = '';

    /**
     * @var array Key-value pairs for report parameters/header info (e.g. User, Date, Filters).
     */
    protected array $parameters = [];

    public function setColumns(array $columns): self
    {
        $this->columns = $columns;
        return $this;
    }

    public function getColumns(): array
    {
        return $this->columns;
    }

    public function setGroupBy(array $groupBy): self
    {
        $this->groupBy = $groupBy;
        return $this;
    }

    public function getGroupBy(): array
    {
        return $this->groupBy;
    }

    public function setTotalizers(array $totalizers): self
    {
        $this->totalizers = $totalizers;
        return $this;
    }

    public function getTotalizers(): array
    {
        return $this->totalizers;
    }

    public function setTitle(string $title): self
    {
        $this->title = $title;
        return $this;
    }

    public function getTitle(): string
    {
        return $this->title;
    }

    public function setParameters(array $parameters): self
    {
        $this->parameters = $parameters;
        return $this;
    }

    public function getParameters(): array
    {
        return $this->parameters;
    }
}
