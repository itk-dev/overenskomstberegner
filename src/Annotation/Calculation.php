<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Annotation;

/**
 * @Annotation
 * @Target({"METHOD"})
 */
class Calculation
{
    /**
     * @required
     *
     * @var string
     */
    public $name;

    /**
     * @Required
     *
     * @var string
     */
    public $description;

    /**
     * @required
     *
     * @var string
     */
    public $formula;

    /**
     * @required
     *
     * @var array
     */
    public $placeholders;

    /**
     * @var string
     */
    public $overenskomsttekst;

    /**
     * @var string
     */
    public $excelFormula;
}
