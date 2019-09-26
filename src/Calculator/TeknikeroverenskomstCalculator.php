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
use App\Annotation\Calculator\Setting;

/**
 * @Calculator(
 *     name="Tekniker og lignende",
 *     description="",
 *     settings={
 *         "overarbejde_fra": @Setting(type="int"),
 *         "overarbejde_til": @Setting(type="int"),
 *     }
 * )
 */
class TeknikeroverenskomstCalculator extends AbstractCalculator
{
    public function calculate(array $input): array
    {
        // TODO: Implement calculate() method.
    }
}
