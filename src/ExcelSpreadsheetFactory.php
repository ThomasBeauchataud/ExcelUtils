<?php

namespace TBCD\Excel;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Helper\Dimension;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * @author Thomas Beauchataud
 * @since 09/11/2021
 */
class ExcelSpreadsheetFactory
{

    private ExcelWorksheetFactory $excelWorksheetFactory;

    /**
     * @param ExcelWorksheetFactory $excelWorksheetFactory
     */
    public function __construct(ExcelWorksheetFactory $excelWorksheetFactory = new ExcelWorksheetFactory())
    {
        $this->excelWorksheetFactory = $excelWorksheetFactory;
    }


    /**
     * @param WorksheetData $worksheetData
     * @return Spreadsheet
     * @throws Exception
     */
    public function createSpreadsheetWithData(WorksheetData $worksheetData): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->removeSheetByIndex(0);
        $worksheet = $this->excelWorksheetFactory->createWorksheetWithData($worksheetData->getData(), $worksheetData->getTitle());

        $spreadsheet->addSheet($worksheet);
        $this->applyDefaultStyles($worksheet, $worksheetData);

        return $spreadsheet;
    }

    /**
     * @param WorksheetData[] $worksheetsData
     * @return Spreadsheet
     * @throws Exception
     */
    public function createSpreadsheetWithMultipleData(array $worksheetsData): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->removeSheetByIndex(0);

        foreach ($worksheetsData as $worksheetData) {
            if (!$worksheetData instanceof WorksheetData) {
                throw new Exception('Objects passed must be instance of ' . WorksheetData::class);
            }
            if (empty($worksheetData->getData())) {
                continue;
            }
            $worksheet = $this->excelWorksheetFactory->createWorksheetWithData($worksheetData->getData(), $worksheetData->getTitle());
            $spreadsheet->addSheet($worksheet);
            $this->applyDefaultStyles($worksheet, $worksheetData);
        }

        if ($spreadsheet->getSheetCount() === 0) {
            $spreadsheet->addSheet(new Worksheet());
        } else {
            $spreadsheet->setActiveSheetIndex(0);
        }

        return $spreadsheet;
    }

    /**
     * @param Worksheet $worksheet
     * @param WorksheetData $worksheetData
     */
    private function applyDefaultStyles(Worksheet $worksheet, WorksheetData $worksheetData): void
    {
        if (empty($worksheetData->getData())) {
            return;
        }

        $alphabet = range("A", "Z");

        $fixedColumnsIndex = [];
        $autoWithColumnsIndex = range(0, count($worksheetData->getData()[0]) - 1);
        $numberFormatColumnsIndex = [];

        foreach ($worksheetData->getData() as $row) {
            $colNumber = 0;
            foreach ($row as $col) {
                if (is_scalar($col) && strlen(strval($col)) > 75 && !in_array($colNumber, $fixedColumnsIndex)) {
                    $fixedColumnsIndex[] = $colNumber;
                    if (in_array($colNumber, $autoWithColumnsIndex)) {
                        unset($autoWithColumnsIndex[array_search($colNumber, $autoWithColumnsIndex)]);
                    }
                }
                if (ctype_digit($col) && !in_array($colNumber, $numberFormatColumnsIndex)) {
                    $numberFormatColumnsIndex[] = $colNumber;
                }
                $colNumber++;
            }
        }

        // Change number format
        foreach ($numberFormatColumnsIndex as $numberFormatColumnIndex) {
            $columnValue = $alphabet[$numberFormatColumnIndex];
            $worksheet->getStyle($columnValue . '1:' . $columnValue . $worksheet->getHighestRow())->getNumberFormat()->setFormatCode('0');
        }

        // Autosize columns
        foreach ($autoWithColumnsIndex as $autoWithColumnIndex) {
            $worksheet->getColumnDimension($alphabet[$autoWithColumnIndex])->setAutoSize(true);
        }

        // Set max size of not autosize columns
        foreach ($fixedColumnsIndex as $fixedColumnIndex) {
            $worksheet->getColumnDimension($alphabet[$fixedColumnIndex])->setWidth(400, Dimension::UOM_PIXELS);
        }

        // Set header to bold
        $worksheet->getStyle('A1:Z1')->getFont()->setBold(true);
    }
}