<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Calculator;

use App\Annotation\Calculation;
use App\Annotation\Calculator;
use App\Calculator\Exception\ValidationException;
use Doctrine\Common\Annotations\AnnotationReader;
use Doctrine\Common\Annotations\AnnotationRegistry;
use Doctrine\Common\Annotations\CachedReader;
use Doctrine\Common\Cache\ArrayCache;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use RuntimeException;

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

    public function getFormulas(AbstractCalculator $calculator)
    {
        AnnotationRegistry::registerLoader('class_exists');
        $reader = new CachedReader(
            new AnnotationReader(),
            new ArrayCache()
        );

        $class = new \ReflectionClass($calculator);
        $calculations = array_filter(array_map(function (\ReflectionMethod $method) use ($reader) {
            $annotation = $reader->getMethodAnnotation($method, Calculation::class);

            return null !== $annotation ? [$method, $annotation] : null;
        }, $class->getMethods()));

        return $calculations;
    }

    /**
     * @param $calculatorClass
     * @param array $settings
     *
     * @return AbstractCalculator
     */
    public function createCalculator($calculatorClass, array $settings)
    {
        $calculator = $this->getCalculator($calculatorClass);
        if (null === $calculator) {
            throw new RuntimeException(sprintf('Invalid calculator: %s', $calculatorClass));
        }

        $settings = $this->normalizeSettings($calculatorClass, $settings);

        return (new $calculatorClass($calculator))->setSettings($settings);
    }

    public function normalizeSettings($calculator, array $values)
    {
        $calculator = $this->getCalculator($calculator);
        if (null !== $calculator) {
            $values = array_filter($values, static function ($name) use ($calculator) {
                return \array_key_exists($name, $calculator['settings']);
            }, ARRAY_FILTER_USE_KEY);

            foreach ($calculator['settings'] as $name => $info) {
                if (!\array_key_exists($name, $values)) {
                    if ($info['required']) {
                        throw new ValidationException(sprintf('Settings %s must be defined.', $name));
                    }
                    if ($info['default']) {
                        $values[$name] = $info['default'];
                    }
                }
                $values[$name] = Calculator::convertToType($name, $info['type'], $values);
            }
        }

        return $values;
    }

    public function calculate(string $calculator, array $settings, array $arguments, $input, array $options = [])
    {
        $calculator = $this->createCalculator($calculator, $settings);
        if (isset($options['test_mode'])) {
            $calculator->setTestMode($options['test_mode']);
        }

        if (\is_string($input)) {
            $input = $this->readInput($input);
        }

        return $calculator->calculate($input, $arguments);
    }

    private function readInput(string $pathname)
    {
        return $this->getInputReader($pathname)->load($pathname);
    }

    private function getInputReader(string $pathname)
    {
        $type = pathinfo($pathname, PATHINFO_EXTENSION);
        switch ($type) {
            case 'csv':
                return new Csv();
            case 'xlsx':
                return new Xlsx();
        }

        throw new RuntimeException(sprintf('Cannot read file of type: %s', $type));
    }
}
