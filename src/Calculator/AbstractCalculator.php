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
use App\Calculator\Exception\InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use ReflectionProperty;

abstract class AbstractCalculator
{
    protected $name;

    /** @var array */
    protected $metadata;

    /** @var array */
    protected $settings;

    /** @var array */
    protected $arguments;

    abstract public function calculate(Spreadsheet $input, array $arguments): Spreadsheet;

    public function __construct(array $metadata)
    {
        $this->metadata = $metadata;
        if (null === $this->name) {
            throw new \RuntimeException(sprintf('Property name must be set on class %s', static::class));
        }
    }

    /**
     * @return mixed
     */
    public function getName()
    {
        return $this->name;
    }

    public function setSettings(array $settings)
    {
        $this->settings = $settings;
        $this->validateAndApplySettings('settings', $settings);

        return $this;
    }

    public function getArguments()
    {
        return $this->metadata['arguments'] ?? [];
    }

    protected function validateAndApplySettings(string $type, array $settings)
    {
        foreach ($this->metadata[$type] as $name => $setting) {
            if ($setting['required']) {
                Calculator::requireValue($name, $settings);
            }
            Calculator::checkType($name, $setting['type'], $settings);
            $settingName = $setting['name'] ?? $name;
            if (!property_exists($this, $name)) {
                throw new InvalidArgumentException(sprintf(
                    'Property "%s" does not exist on %s.',
                    $name,
                    static::class
                ));
            }
            $property = new ReflectionProperty($this, $name);
            $property->setAccessible(true);
            if (\array_key_exists($settingName, $settings)) {
                $property->setValue($this, $settings[$settingName]);
            } elseif (isset($setting['default'])) {
                $property->setValue($this, $setting['default']);
            }
        }
    }
}
