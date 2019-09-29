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
 *     settings={
 *         "min": @Setting(type="int"),
 *         "max": @Setting(type="int", required=false),
 *     },
 *     arguments={
 *         "start": @Argument(type="date"),
 *         "end": @Argument(type="date", required=false),
 *     }
 * )
 */
class StuffCalculator extends AbstractCalculator
{
    /** @var int */
    private $min;

    /** @var int */
    private $max;

    /** @var DateTime */
    private $start;

    /** @var DateTime|null */
    private $end;

    public function calculate(Spreadsheet $input, array $arguments): Spreadsheet
    {
        $this->validateAndApplySettings('arguments', $arguments);

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
