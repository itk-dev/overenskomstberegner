<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\DependencyInjection\Compiler;

use App\Annotation\Calculator;
use App\Calculator\AbstractCalculator;
use App\Calculator\Manager;
use Doctrine\Common\Annotations\AnnotationReader;
use Doctrine\Common\Annotations\AnnotationRegistry;
use Doctrine\Common\Annotations\CachedReader;
use Doctrine\Common\Cache\ArrayCache;
use ReflectionClass;
use Symfony\Component\DependencyInjection\Compiler\CompilerPassInterface;
use Symfony\Component\DependencyInjection\ContainerBuilder;

class AppCompilerPass implements CompilerPassInterface
{
    public function process(ContainerBuilder $container)
    {
        $services = $container->findTaggedServiceIds('app.calculator');
        $calculators = array_filter($services, static function ($class) {
            return is_a($class, AbstractCalculator::class, true);
        }, ARRAY_FILTER_USE_KEY);

        AnnotationRegistry::registerLoader('class_exists');
        $reader = new CachedReader(
            new AnnotationReader(),
            new ArrayCache()
        );

        foreach ($calculators as $class => &$metadata) {
            $reflectionClass = new ReflectionClass($class);
            /** @var Calculator $annotation */
            $annotation = $reader->getClassAnnotation($reflectionClass, Calculator::class);
            $properties = $reflectionClass->getProperties();
            $metadata = [
                'class' => $class,
                'settings' => [],
                'arguments' => [],
            ];
            $metadata += $annotation->asArray();

            foreach ($properties as $property) {
                if (null !== $annotation = $reader->getPropertyAnnotation($property, Calculator\Setting::class)) {
                    $metadata['settings'][$property->getName()] = $annotation->asArray();
                } elseif (null !== $annotation = $reader->getPropertyAnnotation($property, Calculator\Argument::class)) {
                    $metadata['arguments'][$property->getName()] = $annotation->asArray();
                }
            }
        }
        unset($metadata);

        $definition = $container->getDefinition(Manager::class);
        $definition->setArgument('$calculators', $calculators);
    }
}
