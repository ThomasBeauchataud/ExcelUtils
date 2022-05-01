# ExcelUtils

This library to easily create excel worksheets, spreadsheets and files.
>This library is mainly used to send data report.

## Usage
1. Inject the service from container if you have one or create it
```
private ExcelFileFactory $excelFileFactor;

public function __construct(ExcelFileFactory $excelFileFactory) 
{
    $this->excelFileFactory = $excelFileFactory;
}
```
```
$excelWorksheetFactory = new ExcelWorksheetFactory();
$excelSpreadsheetFactory = new ExcelSpreadsheetFactory($worksheetFactory);
$excelFileFactory = new ExcelFileFactory($spreadsheetFactory);
```
2. Create the file with your data
```
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

$filePath = $fileFactory->createFileWithData(new WorksheetData($data, 'my-sheet-name'));
```
3. Get then result below

![alt text](doc/exemple.png)