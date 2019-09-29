<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Entity;

use App\Validator\Constraints\ValidCalculatorSettings;
use Doctrine\ORM\Mapping as ORM;
use Gedmo\Blameable\Traits\BlameableEntity;
use Gedmo\Mapping\Annotation as Gedmo;
use Gedmo\Timestampable\Traits\TimestampableEntity;
use Symfony\Bridge\Doctrine\Validator\Constraints\UniqueEntity;

/**
 * @ORM\Entity(repositoryClass="App\Repository\CalculationRepository")
 * @Gedmo\Loggable
 * @UniqueEntity(fields={"name"})
 */
class Calculation
{
    use BlameableEntity;
    use TimestampableEntity;

    /**
     * @ORM\Id
     * @ORM\GeneratedValue
     * @ORM\Column(type="integer")
     */
    private $id;

    /**
     * @Gedmo\Versioned
     * @ORM\Column(type="string", length=255)
     */
    private $name;

    /**
     * @ORM\Column(type="string", length=255)
     */
    private $calculator;

    /**
     * @ORM\Column(type="json")
     * @ValidCalculatorSettings(calculatorField="calculator")
     */
    private $calculatorSettings = [];

    public function getId(): ?int
    {
        return $this->id;
    }

    public function getName(): ?string
    {
        return $this->name;
    }

    public function setName(string $name): self
    {
        $this->name = $name;

        return $this;
    }

    public function getCalculator(): ?string
    {
        return $this->calculator;
    }

    public function setCalculator(string $calculator): self
    {
        $this->calculator = $calculator;

        return $this;
    }

    public function getCalculatorSettings(): ?array
    {
        return $this->calculatorSettings;
    }

    public function setCalculatorSettings(array $calculatorSettings): self
    {
        $this->calculatorSettings = $calculatorSettings;

        return $this;
    }
}
