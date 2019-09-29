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
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * @Calculator(
 *     name="Tekniker og lignende",
 *     description="",
 *     settings={
 *         "overarbejdeFra": @Setting(type="int", description="Tidspunkt hvor overtidsperioden starter"),
 *         "overarbejdeTil": @Setting(type="int", description="Tidspunkt hvor overtidsperioden slutter"),
 *     },
 *     arguments={
 *         "startTime": @Argument(type="date", name="Start time", description="Start time", required=true),
 *         "endTime": @Argument(type="date", description="End time", required=true),
 *     }
 * )
 */
class TeknikeroverenskomstCalculator extends AbstractCalculator
{
    private $overarbejdeFra;

    private $overarbejdeTil;

    public function calculate(Spreadsheet $input, array $options): Spreadsheet
    {
        // TODO: Implement calculate() method.
    }
}
