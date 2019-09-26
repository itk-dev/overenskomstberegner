<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Calculator;

class Manager
{
    /** @var array */
    private $calculators;

    public function __construct(array $calculators)
    {
        $this->calculators = $calculators;
    }

    /**
     * @return AbstractCalculator[]
     */
    public function getCalculators(): array
    {
        return $this->calculators;
    }

    public function getCalculator($calculator)
    {
        $calculators = $this->getCalculators();

        return $calculators[$calculator] ?? null;
    }
}
