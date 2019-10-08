<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Annotation;

use DateTime;
use InvalidArgumentException;
use Symfony\Component\Form\Extension\Core\Type\CheckboxType;
use Symfony\Component\Form\Extension\Core\Type\DateType;
use Symfony\Component\Form\Extension\Core\Type\IntegerType;
use Symfony\Component\Form\Extension\Core\Type\TextType;

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

    public function asArray(): array
    {
        return [
            'name' => $this->name,
            'description' => $this->description,
        ];
    }

    public static function checkType($name, $typeName, array $values)
    {
        if (isset($values[$name])) {
            $value = $values[$name];
            switch ($typeName) {
                case 'bool':
                    if (!static::isBool($value)) {
                        throw new InvalidTypeException(sprintf('Must be a bool: %s', $name));
                    }
                    break;
                case 'date':
                    if (!static::isDate($value)) {
                        throw new InvalidTypeException(sprintf('Must be a date: %s', $name));
                    }
                    break;
                case 'int':
                    if (!static::isInt($value)) {
                        throw new InvalidTypeException(sprintf('Must be an int: %s', $name));
                    }
                    break;
                case 'string':
                    if (!static::isString($value)) {
                        throw new InvalidTypeException(sprintf('Must be a string: %s', $name));
                    }
                    break;
                default:
                    throw new InvalidTypeException(sprintf('Unknown type: %s', $typeName));
            }

            return $value;
        }

        return null;
    }

    public static function convertToType($name, $typeName, array $values)
    {
        if (isset($values[$name])) {
            $value = $values[$name];
            switch ($typeName) {
                case 'bool':
                    return (bool) $value;
                case 'date':
                    return $value instanceof DateTime ? $value : new DateTime($value);
                case 'int':
                    return (int) $value;
                case 'string':
                    return (string) $value;
                default:
                    throw new InvalidTypeException(sprintf('Unknown type: %s', $typeName));
            }

            return $value;
        }

        return null;
    }

    protected static function isBool($value): bool
    {
        return \is_bool($value);
    }

    protected static function isDate($value): bool
    {
        return $value instanceof DateTime;
    }

    protected static function isInt($value): bool
    {
        return \is_int($value);
    }

    protected static function isString($value): bool
    {
        return \is_string($value);
    }

    public static function requireValue(string $name, array $values): void
    {
        if (!\array_key_exists($name, $values)) {
            throw new InvalidArgumentException(sprintf('Missing value for "%s"', $name));
        }
    }

    public static function getFormType($type)
    {
        if (isset($type['type'])) {
            $type = $type['type'];
        }

        switch ($type) {
            case 'bool':
                return CheckboxType::class;
            case 'date':
                return DateType::class;
            case 'int':
                return IntegerType::class;
        }

        return TextType::class;
    }

    public static function getFormOptions(array $info)
    {
        return [
            'label' => $info['name'],
            'required' => $info['required'],
            'help' => $info['description'] ?? null,
        ];
    }
}
