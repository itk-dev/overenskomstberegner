<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Validator\Constraints;

use App\Calculator\Exception\ValidationException;
use App\Calculator\Manager;
use ReflectionException;
use ReflectionProperty;
use Symfony\Component\Validator\Constraint;
use Symfony\Component\Validator\ConstraintValidator;
use Symfony\Component\Validator\Exception\ConstraintDefinitionException;
use Symfony\Component\Validator\Exception\UnexpectedTypeException;
use Symfony\Component\Validator\Exception\ValidatorException;

class ValidCalculatorSettingsValidator extends ConstraintValidator
{
    /** @var Manager */
    private $manager;

    public function __construct(Manager $manager)
    {
        $this->manager = $manager;
    }

    public function validate($value, Constraint $constraint)
    {
        if (!$constraint instanceof ValidCalculatorSettings) {
            throw new UnexpectedTypeException($constraint, ValidCalculatorSettings::class);
        }

        if (empty($constraint->calculatorField)) {
            throw new ConstraintDefinitionException('calculatorField must be set.');
        }

        $object = $this->context->getObject();
        try {
            $property = new ReflectionProperty($object, $constraint->calculatorField);
            $property->setAccessible(true);
            $calculatorClass = $property->getValue($object);
            try {
                $this->manager->normalizeSettings($calculatorClass, $value);
            } catch (ValidationException $exception) {
                throw new ValidatorException($exception->getMessage());
            }
        } catch (ReflectionException $exception) {
            throw new ConstraintDefinitionException(sprintf('Field %s does not exist', $constraint->calculatorField));
        }
    }
}
