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
use App\Util\DanishHolidays;
use DateTimeImmutable;
use DateTimeInterface;
use PhpOffice\PhpSpreadsheet\Shared\Date as ExcelDate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use ReflectionProperty;

abstract class AbstractCalculator
{
    protected const EXCEL_COLUMN_A = 0;
    protected const EXCEL_COLUMN_B = 1;
    protected const EXCEL_COLUMN_C = 2;
    protected const EXCEL_COLUMN_D = 3;
    protected const EXCEL_COLUMN_E = 4;
    protected const EXCEL_COLUMN_F = 5;
    protected const EXCEL_COLUMN_G = 6;
    protected const EXCEL_COLUMN_H = 7;
    protected const EXCEL_COLUMN_I = 8;
    protected const EXCEL_COLUMN_J = 9;
    protected const EXCEL_COLUMN_K = 10;
    protected const EXCEL_COLUMN_L = 11;
    protected const EXCEL_COLUMN_M = 12;
    protected const EXCEL_COLUMN_N = 13;
    protected const EXCEL_COLUMN_O = 14;
    protected const EXCEL_COLUMN_P = 15;
    protected const EXCEL_COLUMN_Q = 16;
    protected const EXCEL_COLUMN_R = 17;
    protected const EXCEL_COLUMN_S = 18;
    protected const EXCEL_COLUMN_T = 19;
    protected const EXCEL_COLUMN_U = 20;
    protected const EXCEL_COLUMN_V = 21;
    protected const EXCEL_COLUMN_W = 22;
    protected const EXCEL_COLUMN_X = 23;
    protected const EXCEL_COLUMN_Y = 24;
    protected const EXCEL_COLUMN_Z = 25;
    protected const EXCEL_COLUMN_AA = 26;
    protected const EXCEL_COLUMN_AB = 27;
    protected const EXCEL_COLUMN_AC = 28;
    protected const EXCEL_COLUMN_AD = 29;
    protected const EXCEL_COLUMN_AE = 30;
    protected const EXCEL_COLUMN_AF = 31;
    protected const EXCEL_COLUMN_AG = 32;
    protected const EXCEL_COLUMN_AH = 33;
    protected const EXCEL_COLUMN_AI = 34;
    protected const EXCEL_COLUMN_AJ = 35;
    protected const EXCEL_COLUMN_AK = 36;
    protected const EXCEL_COLUMN_AL = 37;
    protected const EXCEL_COLUMN_AM = 38;
    protected const EXCEL_COLUMN_AN = 39;
    protected const EXCEL_COLUMN_AO = 40;
    protected const EXCEL_COLUMN_AP = 41;
    protected const EXCEL_COLUMN_AQ = 42;
    protected const EXCEL_COLUMN_AR = 43;
    protected const EXCEL_COLUMN_AS = 44;
    protected const EXCEL_COLUMN_AT = 45;
    protected const EXCEL_COLUMN_AU = 46;
    protected const EXCEL_COLUMN_AV = 47;
    protected const EXCEL_COLUMN_AW = 48;
    protected const EXCEL_COLUMN_AX = 49;
    protected const EXCEL_COLUMN_AY = 50;
    protected const EXCEL_COLUMN_AZ = 51;

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

    protected function isHoliday($date = null)
    {
        if (null === $date) {
            return false;
        }

        if (is_numeric($date)) {
            $date = $this->getExcelDate($date);
        }

        return DanishHolidays::isHoliday($date);
    }

    protected const WEEKDAY_MONDAY = 1;
    protected const WEEKDAY_TUESDAY = 2;
    protected const WEEKDAY_WEDNESDAY = 3;
    protected const WEEKDAY_THURSDAY = 4;
    protected const WEEKDAY_FRIDAY = 5;
    protected const WEEKDAY_SATURDAY = 6;
    protected const WEEKDAY_SUNDAY = 7;

    /**
     * Get weekday (1: Monday, …, 7: Sunday).
     */
    protected function getWeekday($date)
    {
        if (is_numeric($date)) {
            $date = $this->getExcelDate($date);
        }

        return (int) $date->format('N');
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
