<?php

namespace TBCD\Excel;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\Filesystem\Filesystem;
use Symfony\Component\Filesystem\Path;

/**
 * @author Thomas Beauchataud
 * @since 18/03/2022
 */
class ExcelFileFactory
{

    private ExcelSpreadsheetFactory $excelSpreadsheetFactory;
    private Filesystem $filesystem;

    /**
     * @param ExcelSpreadsheetFactory $excelSpreadsheetFactory
     * @param Filesystem $filesystem
     */
    public function __construct(ExcelSpreadsheetFactory $excelSpreadsheetFactory = new ExcelSpreadsheetFactory(), Filesystem $filesystem = new Filesystem())
    {
        $this->excelSpreadsheetFactory = $excelSpreadsheetFactory;
        $this->filesystem = $filesystem;
    }


    /**
     * @param WorksheetData $worksheetData
     * @param string $filePath
     * @return string
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createFileWithData(WorksheetData $worksheetData, string $filePath): string
    {
        $spreadsheet = $this->excelSpreadsheetFactory->createSpreadsheetWithData($worksheetData);
        return $this->createFile($spreadsheet, $filePath);
    }

    /**
     * @param WorksheetData[] $worksheetsData
     * @param string $filePath
     * @return string
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createFileWithMultipleData(array $worksheetsData, string $filePath): string
    {
        $spreadsheet = $this->excelSpreadsheetFactory->createSpreadsheetWithMultipleData($worksheetsData);
        return $this->createFile($spreadsheet, $filePath);
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string $filePath
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function createFile(Spreadsheet $spreadsheet, string $filePath): string
    {
        if (Path::getExtension($filePath) !== 'xlsx' && Path::getExtension($filePath) !== 'xls') {
            $filePath = Path::changeExtension($filePath, 'xlsx');
        }

        if (($directory = Path::getDirectory($filePath)) !== "" && !$this->filesystem->exists($directory)) {
            $this->filesystem->mkdir($directory);
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($filePath);
        return $filePath;
    }
}