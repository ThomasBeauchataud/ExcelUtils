<?php

namespace TBCD\Excel;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\DependencyInjection\ParameterBag\ParameterBagInterface;
use Symfony\Component\Filesystem\Filesystem;

/**
 * @author Thomas Beauchataud
 * @since 18/03/2022
 */
class ExcelFileFactory
{

    private ExcelSpreadsheetFactory $excelSpreadsheetFactory;
    private Filesystem $filesystem;
    private string $workspace = '.';

    /**
     * @param ExcelSpreadsheetFactory $excelSpreadsheetFactory
     * @param ParameterBagInterface|null $parameterBag
     */
    public function __construct(ExcelSpreadsheetFactory $excelSpreadsheetFactory, ParameterBagInterface $parameterBag = null)
    {
        $this->excelSpreadsheetFactory = $excelSpreadsheetFactory;
        $this->filesystem = new Filesystem();
        if ($parameterBag && $parameterBag->has('kernel.project_dir')) {
            $this->workspace = $parameterBag->get('kernel.project_dir') . "\\var\\tmp\\";
        }
    }


    /**
     * @param WorksheetData $worksheetData
     * @param string|null $fileName
     * @param string|null $directory
     * @return string
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createFileWithData(WorksheetData $worksheetData, string $fileName = null, string $directory = null): string
    {
        $spreadsheet = $this->excelSpreadsheetFactory->createSpreadsheetWithData($worksheetData);
        return $this->createFile($spreadsheet, $fileName, $directory);
    }

    /**
     * @param WorksheetData[] $worksheetsData
     * @param string|null $fileName
     * @param string|null $directory
     * @return string
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createFileWithMultipleData(array $worksheetsData, string $fileName = null, string $directory = null): string
    {
        $spreadsheet = $this->excelSpreadsheetFactory->createSpreadsheetWithMultipleData($worksheetsData);
        return $this->createFile($spreadsheet, $fileName, $directory);
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string|null $fileName
     * @param string|null $directory
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function createFile(Spreadsheet $spreadsheet, string $fileName = null, string $directory = null): string
    {
        $fileName = $fileName ?? (uniqid() . '.xlsx');
        $directory = $directory ?? $this->workspace;
        $writer = new Xlsx($spreadsheet);
        if (!$this->filesystem->exists($directory)) {
            $this->filesystem->mkdir($directory);
        }
        $filePath = $directory . $fileName;
        $writer->save($filePath);
        return $filePath;
    }
}