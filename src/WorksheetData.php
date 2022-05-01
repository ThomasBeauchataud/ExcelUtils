<?php

namespace TBCD\Excel;

/**
 * @author Thomas Beauchataud
 * @since 15/11/2021
 */
class WorksheetData
{

    /**
     * @var array
     */
    private array $data;

    /**
     * @var string
     */
    private string $title;

    /**
     * @param array $data
     * @param string $title
     */
    public function __construct(array $data, string $title)
    {
        $this->data = $data;
        $this->title = $title;
    }


    /**
     * @return array
     */
    public function getData(): array
    {
        return $this->data;
    }

    /**
     * @return string
     */
    public function getTitle(): string
    {
        return $this->title;
    }
}