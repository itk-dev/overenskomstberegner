<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Annotation;

use App\Annotation\Calculator\Setting;

/**
 * @Annotation
 * @Target({"CLASS"})
 */
class Calculator
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
     * @var array
     */
    public $settings;

    /**
     * @required
     *
     * @var array
     */
    public $arguments;

    public function asArray(): array
    {
        return [
            'name' => $this->name,
            'description' => $this->description,
            'settings' => array_map(static function (Setting $setting) { return $setting->asArray(); }, $this->settings),
        ];
    }
}
