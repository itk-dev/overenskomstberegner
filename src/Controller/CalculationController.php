<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Controller;

use App\Annotation\Calculator;
use App\Calculator\Manager;
use App\Entity\Calculation;
use EasyCorp\Bundle\EasyAdminBundle\Controller\EasyAdminController;
use Symfony\Component\Form\Extension\Core\Type\FormType;
use Symfony\Component\Form\Extension\Core\Type\IntegerType;
use Symfony\Component\Form\Extension\Core\Type\TextType;

class CalculationController extends EasyAdminController
{
    /** @var Manager */
    private $manager;

    public function __construct(Manager $manager)
    {
        $this->manager = $manager;
    }

    protected function createEntityFormBuilder($entity, $view)
    {
        $builder = parent::createEntityFormBuilder($entity, $view);

        $settingsBuilder = $builder->create('calculatorSettings', FormType::class, [
            'compound' => true,
            'attr' => [
                'class' => 'calculator-settings',
            ],
        ]);
        foreach ($this->manager->getCalculators() as $class => $calculator) {
            $formName = str_replace('\\', '_', $class);
            $bb = $settingsBuilder->create($formName, FormType::class, [
                'label' => $calculator['name'] ?? $class,
            ]);
            $this->buildCalculatorForm($calculator, $bb);
            $settingsBuilder->add($bb);
        }
        $builder->add($settingsBuilder);

        return $builder;
    }

    private function buildCalculatorForm(array $calculator, $builder)
    {
        foreach ($calculator['settings'] as $name => $info) {
            $builder->add(
//                $this->getCalculatorSettingFormName($calculator, $name),
                $name,
                $this->getFormType($info['type']),
                [
                    'mapped' => false,
                    'label' => $info['name'] ?? $name,
                    'required' => $info['required'],
                    'help' => $info['description'] ?? null,
                    'attr' => [
                        'data-calculator' => $calculator['class'],
                    ],
                ]
            );
        }
    }

    private function getCalculatorSettingFormName($calculator, $name)
    {
        return $this->getCalculatorId($calculator).'-'.$name;
    }

    private function getCalculatorSettingName($calculator, $name)
    {
        $prefix = $this->getCalculatorId($calculator).'-';

        return 0 === strpos($name, $prefix) ? substr($name, \strlen($prefix)) : null;
    }

    private function getCalculatorId($calculator)
    {
        if ($calculator instanceof Calculator) {
            $calculator = \get_class($calculator);
        }

        if (\is_array($calculator) && isset($calculator['class'])) {
            $calculator = $calculator['class'];
        }

        return md5($calculator);
    }

    private function getFormType(string $type)
    {
        switch ($type) {
            case 'int':
                return IntegerType::class;
        }

        return TextType::class;
    }

    protected function persistEntity($entity)
    {
        return parent::persistEntity($entity);
    }

    protected function updateEntity($entity)
    {
        $editForm = \func_num_args() > 1 ? func_get_arg(1) : null;

        if ($entity instanceof Calculation) {
            $calculator = $this->manager->getCalculator($entity->getCalculator());

            $builder = $this->createFormBuilder();
            $this->buildCalculatorForm($calculator['class'], $calculator['settings'], $builder);
            $form = $builder->getForm();
            $form->handleRequest($this->request);

            $calculatorSettings = [];
            $values = $this->request->request->get('calculation');
            foreach ($values as $name => $value) {
                if (null !== $name = $this->getCalculatorSettingName($calculator, $name)) {
                    $calculatorSettings[$name] = $value;
                }
            }
            $entity->setCalculatorSettings($calculatorSettings);
        }

        return parent::updateEntity($entity);
    }
}
