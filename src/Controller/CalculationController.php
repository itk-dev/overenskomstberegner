<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Controller;

use Symfony\Component\HttpFoundation\RedirectResponse;
use App\Annotation\Calculator;
use App\Calculator\Manager;
use App\Entity\Calculation;
use EasyCorp\Bundle\EasyAdminBundle\Controller\EasyAdminController;
use RuntimeException;
use Symfony\Component\Form\Extension\Core\Type\FormType;
use Symfony\Component\Form\FormBuilderInterface;
use Symfony\Component\Form\FormEvent;
use Symfony\Component\Form\FormEvents;

class CalculationController extends EasyAdminController
{
    /** @var Manager */
    private $manager;

    public function __construct(Manager $manager)
    {
        $this->manager = $manager;
    }

    public function runAction(): RedirectResponse
    {
        $id = $this->request->get('id');

        return $this->redirectToRoute('calculate_run', ['id' => $id]);
    }

    protected function createEntityFormBuilder($entity, $view)
    {
        if (!$entity instanceof Calculation) {
            throw new RuntimeException(sprintf('Invalid entity: %s', \get_class($entity)));
        }

        $data = $this->buildData($entity);

        $builder = parent::createEntityFormBuilder($entity, $view);

        $settingsBuilder = $builder->create('calculatorSettings', FormType::class, [
            'compound' => true,
            'attr' => [
                'class' => 'calculator-settings-forms',
            ],
            // We'll inject new fields into this form.
            'allow_extra_fields' => true,
        ])->addEventListener(FormEvents::PRE_SUBMIT, function (FormEvent $event) {
            $data = $event->getData();
            $form = $event->getForm();
            $calculator = $this->request->request->get($form->getParent()->getName())['calculator'] ?? null;
            $formName = $this->getFormName($calculator);
            if (isset($data[$formName])) {
                foreach ($form->getData() as $name => $value) {
                    $form->remove($name);
                }
                $form->setData($this->manager->normalizeSettings($calculator, $data[$formName]));
            }
        });
        foreach ($this->manager->getCalculators() as $class => $calculator) {
            $formName = $this->getFormName($calculator);
            $calculatorSettingsBuilder = $settingsBuilder->create($formName, FormType::class, [
                'label' => $calculator['name'] ?? $class,
                'attr' => [
                    'class' => 'calculator-settings',
                    'data-calculator' => $class,
                ],
            ]);
            $this->buildCalculatorForm($calculator, $calculatorSettingsBuilder, $data[$formName]);
            $settingsBuilder->add($calculatorSettingsBuilder);
        }
        $builder->add($settingsBuilder);

        return $builder;
    }

    private function getFormName($class)
    {
        if (isset($class['class'])) {
            $class = $class['class'];
        }

        return str_replace('\\', '_', $class);
    }

    private function buildData(Calculation $calculation)
    {
        $data = [];
        foreach ($this->manager->getCalculators() as $calculator) {
            $formName = $this->getFormName($calculator);
            $data[$formName] = $calculation->getCalculator() === $calculator['class'] ? $calculation->getCalculatorSettings() : [];
        }

        return $data;
    }

    private function buildCalculatorForm(array $calculator, FormBuilderInterface $builder, array $data)
    {
        foreach ($calculator['settings'] as $name => $info) {
            $builder->add(
                $name,
                Calculator::getFormType($info),
                Calculator::getFormOptions($info)
                    + [
                        'attr' => [
                            'data-calculator' => $calculator['class'],
                        ],
                        'data' => $data[$name] ?? null,
                    ]
            );
        }
    }
}
