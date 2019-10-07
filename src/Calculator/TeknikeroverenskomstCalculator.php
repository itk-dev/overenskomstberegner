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
    protected $name = 'Overenskomst for teknikere';

    private $overarbejdeFra;

    private $overarbejdeTil;

    public function calculate(Spreadsheet $input, array $options): Spreadsheet
    {
        // 1. Read input and group by employee and sort by dates.
        // 2. Calculate sums of TF.
        // 3. Generate output: One line per TF per employee.

        throw new \RuntimeException(__METHOD__.' not implemented!');
    }

    /**
     * Calculate arbejdstimer.
     *
     * @Calculation(
     *     overenskomsttekst="…",
     *     excel_formula="=HVIS(ELLER(F3="Vagt";OG(F3="Sygdom";J3<>0));J3-I3;HVIS(ER.FEJL(LOPSLAG(F3;Normnedsættende;1;FALSK));0;LOPSLAG(C3;Meta!H:I;2;FALSK)/5/24))",
     *     calculation="\frac{a}{b}",
     *     placeholders={
     *         "a": @Placeholder(name="The a value", description="Value of a", type="int"),
     *         "b": @Placeholder(name="The a value", description="Value of a", type="int"),
     *     }
     * )
     */
    private function calculateArbejdstimer()
    {
    }
}
