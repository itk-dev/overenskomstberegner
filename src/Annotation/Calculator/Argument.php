<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Annotation\Calculator;

use Doctrine\Common\Annotations\Annotation\Target;

/**
 * @Annotation
 * @Target("ANNOTATION")
 */
class Argument
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
    public $description;

    /**
     * @required
     *
     * @Enum({"string", "bool", "int", "date"})
     *
     * @var string
     */
    public $type;

    /**
     * @var bool
     */
    public $required = true;

    /**
     * Default value.
     *
     * @var mixed
     */
    public $default;

    public function asArray(): array
    {
        return [
            'name' => $this->name,
            'description' => $this->description,
            'type' => $this->type,
            'required' => $this->required,
            'default' => $this->default,
        ];
    }
}
