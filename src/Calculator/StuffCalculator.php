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
 *     name="Stuff",
 *     description="",
 *     settings={
 *         "min": @Setting(type="int"),
 *         "max": @Setting(type="int", required=false),
 *     }
 * )
 */
class StuffCalculator extends AbstractCalculator
{
    public function calculate(array $input): array
    {
        // TODO: Implement calculate() method.
    }
}
