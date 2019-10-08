<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Calculator;

use App\Annotation\Calculator;
use App\Annotation\Calculator\Argument;
use App\Annotation\Calculator\Setting;
use DateTime;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * @Calculator(
 *     name="Stuff",
 *     description="",
 * )
 */
class StuffCalculator extends AbstractCalculator
{
    /**
     * @Setting(type="int"),
     *
     * @var int */
    private $min;

    /**
     * @Setting(type="int", required=false),
     *
     * @var int
     */
    private $max;

    /**
     * @Argument(type="date"),
     *
     * @var DateTime
     */
    private $start;

    /**
     * @Argument(type="date", required=false),
     *
     * @var DateTime|null
     */
    private $end;

    protected function load(Spreadsheet $input): void
    {
        // TODO: Implement load() method.
    }

    protected function doCalculate(): void
    {
        // TODO: Implement doCalculate() method.
    }

    protected function render(): Spreadsheet
    {
        $result = new Spreadsheet();
        $sheet = $result->getActiveSheet();
        $sheet->setCellValueByColumnAndRow(1, 1, $this->start->format('Y-m-d'));

        $row = 2;
        for ($i = $this->min; $i <= $this->max; ++$i) {
            $sheet
                ->setCellValueByColumnAndRow(1, $row, $i)
                ->setCellValueByColumnAndRow(2, $row, 2 * $i)
                ->setCellValueByColumnAndRow(3, $row, '=42*'.$sheet->getCellByColumnAndRow(2, $row)->getCoordinate());
            ++$row;
        }

        return $result;
    }
}
