<?php

/*
 * The file is part of the WoWUltimate project 
 * 
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 *
 * Author Thomas Beauchataud
 * From 01/05/2022
 */

use PHPUnit\Framework\TestCase;
use Symfony\Component\Filesystem\Filesystem;
use TBCD\Excel\ExcelFileFactory;
use TBCD\Excel\ExcelSpreadsheetFactory;
use TBCD\Excel\ExcelWorksheetFactory;
use TBCD\Excel\WorksheetData;

class Test extends TestCase
{

    /**
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function testAll(): void
    {
        $worksheetFactory = new ExcelWorksheetFactory();
        $spreadsheetFactory = new ExcelSpreadsheetFactory($worksheetFactory);
        $fileFactory = new ExcelFileFactory($spreadsheetFactory);
        $filesystem = new Filesystem();

        $data = [
            [
                'column1' => 'row1',
                'column2' => 'row1',
                'column3' => 'row1'
            ],
            [
                'column1' => 'row2',
                'column2' => 'row2',
                'column3' => 'row2'
            ]
        ];

        $filePath = $fileFactory->createFileWithData(new WorksheetData($data, 'test'));
        $this->assertTrue($filesystem->exists($filePath));
        $filesystem->remove($filePath);
    }
}