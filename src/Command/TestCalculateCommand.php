<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Command;

use App\Calculator\Manager;
use App\Entity\Calculation;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Helper\Table;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Yaml\Yaml;

class TestCalculateCommand extends Command
{
    protected static $defaultName = 'app:test:calculate';

    /** @var Manager */
    private $manager;

    public function __construct(Manager $manager)
    {
        parent::__construct();
        $this->manager = $manager;
    }

    protected function configure()
    {
        $this->addArgument('pattern', InputArgument::OPTIONAL);
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $testsDir = realpath(__DIR__.'/../../tests/calculations');
        $filenames = glob($testsDir.'/*.yaml');
        foreach ($filenames as $filename) {
            if ((null !== $pattern = $input->getArgument('pattern'))
                && !fnmatch($pattern, $filename)) {
                continue;
            }
            $output->writeln($filename);

            $yaml = file_get_contents($filename);
            $config = Yaml::parse($yaml);

            $source = __DIR__.'/../../tests/input/'.$config['input'];
            $calculation = (new Calculation())
                ->setCalculator($config['calculator'])
                ->setCalculatorSettings($config['settings']);

            $arguments = $config['arguments'];

            $output->writeln($source);

            $result = $this->manager->calculate(
                $calculation->getCalculator(),
                $calculation->getCalculatorSettings(),
                $arguments,
                $source,
                [
                    'test_mode' => true,
                ]
            );

            $data = $result->getSheet(0)->toArray();
            $table = new Table($output);
            foreach ($data as $row) {
                $table->addRow($row);
            }
            $table->render();
        }
    }
}
