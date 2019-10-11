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
use DateTimeImmutable;
use DateTimeInterface;
use PhpOffice\PhpSpreadsheet\Shared\Date as ExcelDate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
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

    public function calculate(Spreadsheet $input, array $arguments): Spreadsheet
    {
        $this->load($input);
        $this->validateAndApplySettings('arguments', $arguments);
        $this->doCalculate();

        return $this->render();
    }

    /**
     * Load data from spreadsheet.
     *
     * @param Spreadsheet $input
     */
    abstract protected function load(Spreadsheet $input): void;

    /**
     * Perform actual calculation.
     */
    abstract protected function doCalculate(): void;

    /**
     * Render calculation result in a spreadsheet.
     *
     * @return Spreadsheet
     */
    abstract protected function render(): Spreadsheet;

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
            if (!property_exists($this, $name)) {
                throw new InvalidArgumentException(sprintf(
                    'Property "%s" does not exist on %s.',
                    $name,
                    static::class
                ));
            }
            $property = new ReflectionProperty($this, $name);
            $property->setAccessible(true);
            if (\array_key_exists($name, $settings)) {
                $property->setValue($this, $settings[$name]);
            } elseif (isset($setting['default'])) {
                $property->setValue($this, $setting['default']);
            }
        }
    }

    /**
     * @param $value
     *
     * @return DateTimeImmutable|null
     *
     * @throws \Exception
     */
    protected function getExcelDate($value)
    {
        return $value ? Date::createFromDateTime(DateTimeImmutable::createFromMutable(ExcelDate::excelToDateTimeObject($value))) : null;
    }

    protected function formatExcelDate(float $excelTimestamp = null)
    {
        return null !== $excelTimestamp ? ExcelDate::excelToDateTimeObject($excelTimestamp)->format('Y-m-d') : '';
    }

    protected function formatExcelDateTime(float $excelTimestamp = null)
    {
        return null !== $excelTimestamp ? ExcelDate::excelToDateTimeObject($excelTimestamp)->format('Y-m-d H:i:s') : '';
    }

    protected function formatExcelTime(float $excelTimestamp = null)
    {
        return null !== $excelTimestamp ? ExcelDate::excelToDateTimeObject($excelTimestamp)->format('H:i') : '';
    }

    protected function dateTime2Excel(DateTimeInterface $dateTime)
    {
        return ExcelDate::PHPToExcel($dateTime);
    }

    protected function time2Excel(DateTimeInterface $date)
    {
        return ExcelDate::formattedPHPToExcel(
            1900,
            1,
            0,
            (int) $date->format('H'),
            (int) $date->format('i'),
            (int) $date->format('s')
        );
    }

    protected function writeCells(Worksheet $spreadsheet, int $columnIndex, int $row, array $cells): void
    {
        foreach ($cells as $cell) {
            $spreadsheet->setCellValueByColumnAndRow($columnIndex, $row, $cell);
            ++$columnIndex;
        }
    }

    protected function writeCell(Worksheet $spreadsheet, int $columnIndex, int $row, $value, $colSpan = 1): void
    {
        $cells = array_fill(0, $colSpan, 'null');
        $cells[0] = $value;
        $this->writeCells($spreadsheet, $columnIndex, $row, $cells);
        if ($colSpan > 1) {
            $spreadsheet->mergeCellsByColumnAndRow($columnIndex, $row, $columnIndex + $colSpan - 1, 1);
        }
    }
}
