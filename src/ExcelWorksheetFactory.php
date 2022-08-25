<?php

namespace TBCD\Excel;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * @author Thomas Beauchataud
 * @since 15/11/2021
 */
class ExcelWorksheetFactory
{

    /**
     * Create a worksheet with data from an associative array and the worksheet name
     *
     * @param array $data
     * @param string $title
     * @return Worksheet
     */
    public function createWorksheetWithData(array $data, string $title): Worksheet
    {
        $worksheet = new Worksheet();
        $worksheet->setTitle($title);

        // Alphabet for headers creation
        $alphabet1[0] = 0;
        $alphabet2 = range("A", "Z");
        $aAlphabet = array_merge($alphabet1, $alphabet2);

        if (empty($data)) {
            return $worksheet;
        }

        $headers = array_keys($data[0]);
        foreach ($headers as $i => $header) {
            $worksheet->setCellValue($aAlphabet[$i + 1] . '1', $header);
        }

        $rows = array_values($data);
        foreach ($rows as $i => $row) {
            $i = $i + 2;
            $row = array_values($row);
            foreach ($row as $k => $column) {
                $k = $k + 1;
                if ($column instanceof \BackedEnum) {
                    $column = $column->value;
                }
                $worksheet->setCellValue($aAlphabet[$k] . $i, $column);
            }
        }

        return $worksheet;
    }
}

