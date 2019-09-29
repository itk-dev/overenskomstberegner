<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Validator\Constraints;

use Symfony\Component\Validator\Constraint;

/**
 * @Annotation
 */
class ValidCalculatorSettings extends Constraint
{
    public $calculatorField;

    public function validatedBy()
    {
        return \get_class($this).'Validator';
    }

    public function getDefaultOption()
    {
        return 'calculatorField';
    }

    public function getRequiredOptions()
    {
        return ['calculatorField'];
    }
}
