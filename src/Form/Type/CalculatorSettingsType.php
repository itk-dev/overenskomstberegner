<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Form\Type;

use App\Calculator\Manager;
use Symfony\Component\Form\ChoiceList\Loader\CallbackChoiceLoader;
use Symfony\Component\Form\Extension\Core\Type\FormType;
use Symfony\Component\Form\FormBuilderInterface;

class CalculatorSettingsType extends FormType
{
    /** @var Manager */
    private $manager;

    public function __construct(Manager $manager)
    {
        parent::__construct();
        $this->manager = $manager;
    }

    public function buildForm(FormBuilderInterface $builder, array $options)
    {
        $options['choice_loader'] = new CallbackChoiceLoader(function () {
            $options = [];

            foreach ($this->manager->getCalculators() as $class => $calculator) {
                $options[$calculator['name']] = $class;
            }

            return $options;
        });
        parent::buildForm($builder, $options);
    }
}
