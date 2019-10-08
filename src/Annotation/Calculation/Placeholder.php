<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Annotation\Calculation;

/**
 * @Annotation
 * @Target({"ANNOTATION"})
 */
class Placeholder
{
    /**
     * @required
     *
     * @var string
     */
    public $name;

    /**
     * @required
     *
     * @var string
     */
    public $type;

    /**
     * @required
     *
     * @var string
     */
    public $description;
}
