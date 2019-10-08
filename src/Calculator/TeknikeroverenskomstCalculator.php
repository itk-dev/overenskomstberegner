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
use App\Annotation\Calculation\Placeholder;
use App\Annotation\Calculator;
use App\Annotation\Calculator\Argument;
use App\Annotation\Calculator\Setting;
use DateTime;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * @Calculator(
 *     name="Tekniker og lignende",
 *     description=""
 * )
 */
class TeknikeroverenskomstCalculator extends AbstractCalculator
{
    protected $name = 'Overenskomst for teknikere';

    /**
     * @Setting(type="string", description="Overskrift på resultat", name="hest"),
     *
     * @var string
     */
    private $resultTitle;

    /**
     * @Setting(type="int", description="Tidspunkt hvor overtidsperioden starter"),
     *
     * @var int
     */
    private $overarbejdeFra;

    /**
     * @Setting(type="int", description="Tidspunkt hvor overtidsperioden slutter"),
     *
     * @var int
     */
    private $overarbejdeTil;

    /**
     * @Argument(type="date", name="Start time", description="Start time", required=true),
     *
     * @var DateTime
     */
    private $startDate;

    /**
     * @Argument(type="date", name="End time", description="End time", required=true),
     *
     * @var DateTime
     */
    private $endDate;

    private const COLUMN_INPUT_NAME = 'name';
    private const COLUMN_INPUT_EMPLOYEE_NUMBER = 'employee number';
    private const COLUMN_INPUT_CONTRACT = 'contract';
    private const COLUMN_INPUT_EMAIL = 'email';
    private const COLUMN_INPUT_DATE = 'date';
    private const COLUMN_INPUT_EVENT = 'event';
    private const COLUMN_INPUT_PLANNED_START = 'planned start';
    private const COLUMN_INPUT_PLANNED_END = 'planned end';
    private const COLUMN_INPUT_ACTUAL_START = 'actual start';
    private const COLUMN_INPUT_ACTUAL_END = 'actual end';

    private const COLUMN_OUTPUT_EMPLOYEE_NUMBER = 'Medarbejdernummer';
    private const COLUMN_OUTPUT_LOENART = 'Lønart';
    private const COLUMN_OUTPUT_LOEBENR = 'Løbenr.';
    private const COLUMN_OUTPUT_ENHEDER_I_ALT = 'Enheder (i alt)';
    private const COLUMN_OUTPUT_IKRAFT_DATO = 'Ikraft dato (for lønmåned)';
    private const COLUMN_OUTPUT_IS_OVERTIME = 'is overtime';

    /**
     * Data grouped by employee number and sorted ascending by date.
     *
     * @var array
     */
    private $data;

    private $result;

    /**
     * Read input and group by employee and sort ascending by dates.
     *
     * {@inheritdoc}
     */
    protected function load(Spreadsheet $spreadsheet): void
    {
        // Get the first sheet.
        $sheet = $spreadsheet->getSheet(0);

        $dataColumnStart = 'A';
        $dataColumnEnd = 'J';
        $dataRowStart = 3;
        $dataRowEnd = $sheet->getHighestRow('E');
        $dataRange = $dataColumnStart.$dataRowStart.':'.$dataColumnEnd.$dataRowEnd;

        $rows = $sheet->rangeToArray($dataRange, null, false, false);

        // Convert values to something useful and group by employee number.
        $this->data = [];
        foreach ($rows as $row) {
            $employeeNumber = (string) $row[1];
            if (!empty($employeeNumber)) {
                $employeeRow = [
                    self::COLUMN_INPUT_NAME => $row[0],
                    self::COLUMN_INPUT_EMPLOYEE_NUMBER => (string) $row[1],
                    self::COLUMN_INPUT_CONTRACT => $row[2],
                    self::COLUMN_INPUT_EMAIL => $row[3],
                    self::COLUMN_INPUT_DATE => $this->getExcelDate($row[4]),
                    self::COLUMN_INPUT_EVENT => $row[5],
                    self::COLUMN_INPUT_PLANNED_START => $this->getExcelTime($row[6]),
                    self::COLUMN_INPUT_PLANNED_END => $this->getExcelTime($row[7]),
                    self::COLUMN_INPUT_ACTUAL_START => $this->getExcelTime($row[8]),
                    self::COLUMN_INPUT_ACTUAL_END => $this->getExcelTime($row[9]),
                ];
                // Assume that actual end is on next day is less that actual start.
                if ($employeeRow[self::COLUMN_INPUT_ACTUAL_END] < $employeeRow[self::COLUMN_INPUT_ACTUAL_START]) {
                    $employeeRow[self::COLUMN_INPUT_ACTUAL_END]->add(new \DateInterval('P1D'));
                }
                $this->data[$employeeNumber][] = $employeeRow;
            }
        }

        // Sort each employee's data by date (ascending).
        foreach ($this->data as &$items) {
            usort($items, static function ($a, $b) {
                return $a[self::COLUMN_INPUT_DATE] <=> $b[self::COLUMN_INPUT_DATE];
            });
        }
    }

    /**
     * Calculate sums of TF.
     */
    protected function doCalculate(): void
    {
        $this->result = [];

        foreach ($this->data as $employeeNumber => $rows) {
            $this->calculateIsOvertime($rows);
            $this->setDates($rows);
            $this->result[$employeeNumber] = $this->calculateEmployee($employeeNumber, $rows);
        }
    }

    /**
     * Generate output: One line per TF per employee.
     *
     * @return Spreadsheet
     */
    protected function render(): Spreadsheet
    {
        $result = new Spreadsheet();
        $sheet = $result->getActiveSheet();
        $rowIndex = 1;
        $this->writeCell($sheet, 1, 1, $this->resultTitle, 5);
        ++$rowIndex;

        $this->writeCells($sheet, 1, $rowIndex, [
            self::COLUMN_OUTPUT_EMPLOYEE_NUMBER,
            self::COLUMN_OUTPUT_LOENART,
            self::COLUMN_OUTPUT_LOEBENR,
            self::COLUMN_OUTPUT_ENHEDER_I_ALT,
            self::COLUMN_OUTPUT_IKRAFT_DATO,
        ]);
        ++$rowIndex;

//        $sheet->setCellValueByColumnAndRow(2, $rowIndex, $this->startDate->format('Y-m-d'));
//        $sheet->setCellValueByColumnAndRow(3, $rowIndex, $this->endDate->format('Y-m-d'));
//        ++$rowIndex;

        foreach ($this->result as $employeeNumber => $row) {
            $columnIndex = 1;
            foreach ($row as $cell) {
                $sheet->setCellValueByColumnAndRow($columnIndex, $rowIndex, $cell);
                ++$columnIndex;
            }
            ++$rowIndex;
        }

        return $result;
    }

    private function calculateEmployee(string $employeeNumber, array $rows)
    {
        foreach ($rows as $index => $row) {
            $row['timer'] = $this->calculateTimer($row);
            $row['overtid'] = $this->calculateOvertid($row, $index, $rows);
        }

        return [
            self::COLUMN_OUTPUT_EMPLOYEE_NUMBER => $employeeNumber,
            self::COLUMN_OUTPUT_LOENART => null,
            self::COLUMN_OUTPUT_LOEBENR => null,
            self::COLUMN_OUTPUT_ENHEDER_I_ALT => null,
            self::COLUMN_OUTPUT_IKRAFT_DATO => null,
        ];
    }

    // @TODO: Is this the complete list?
    private const EVENT_VAGT = 'Vagt';
    private const EVENT_SYGDOM = 'Sygdom';
    private const EVENT_LØN_OVERTID = 'Løn: Overtid';
    private const EVENT_LØN_IKKE_PLANLAGT_7_DAG = 'Løn: Ikke planlagt/7. dag';
    private const EVENT_KURSUS = 'Kursus';
    private const EVENT_FERIETIMER = 'Ferietimer';
    private const EVENT_SENIORDAG = 'Seniordag';
    private const EVENT_LØN_DELT_TJENESTE = 'Løn: Delt tjeneste';

    // @TODO: Is this the complete list?
    private const CONTRACT_TEKNIK_37_HOURS_3_MÅNEDER = 'Teknik 37 hours 3 måneder';
    private const CONTRACT_TEKNIK_37_HOURS = 'Teknik 37 hours';
    private const CONTRACT_TEKNIK_32_HOURS = 'Teknik 32 hours';
    private const CONTRACT_TIMELØNNEDE = 'Timelønnede';

    /**
     * Calculate arbejdstimer.
     *
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="\frac{overarbejdeFra}{overarbejdeTil}",
     *     placeholders={
     *         "a": @Placeholder(name="The a value", description="Value of a", type="int"),
     *         "b": @Placeholder(name="The a value", description="Value of a", type="int"),
     *     },
     *     overenskomsttekst="…",
     *     excelFormula="=HVIS(ELLER(F3=""Vagt"";OG(F3=""Sygdom"";J3<>0));J3-I3;HVIS(ER.FEJL(LOPSLAG(F3;Normnedsættende;1;FALSK));0;LOPSLAG(C3;Meta!H:I;2;FALSK)/5/24))",
     * )
     */
    private function calculateTimer(array $row)
    {
        if (self::EVENT_VAGT === $row[self::COLUMN_INPUT_EVENT]
            || (self::EVENT_SYGDOM === $row[self::COLUMN_INPUT_EVENT] && !empty($row[self::COLUMN_INPUT_ACTUAL_END]))) {
            $interval = $row[self::COLUMN_INPUT_ACTUAL_END]->diff($row[self::COLUMN_INPUT_ACTUAL_START]);
            $hours = $interval->h;
            if ($interval->i > 0) {
                $hours += $interval->i * (100 / 60);
            }

            return $hours;
        }
    }

    /**
     * Calculate if a row is overtime.
     *
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculateIsOvertime(array &$rows)
    {
        $numberOfRows = \count($rows);
        foreach ($rows as $index => &$row) {
            $row[self::COLUMN_OUTPUT_IS_OVERTIME] =
                ($index > 0 && $row[self::COLUMN_INPUT_DATE] === $rows[$index - 1][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index - 1][self::COLUMN_INPUT_EVENT])
                || ($index > 1 && $row[self::COLUMN_INPUT_DATE] === $rows[$index - 2][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index - 2][self::COLUMN_INPUT_EVENT])
                || ($index < $numberOfRows - 1 && $row[self::COLUMN_INPUT_DATE] === $rows[$index + 1][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index + 1][self::COLUMN_INPUT_EVENT])
                || ($index < $numberOfRows - 2 && $row[self::COLUMN_INPUT_DATE] === $rows[$index + 2][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index + 2][self::COLUMN_INPUT_EVENT]);
        }
    }

    /**
     * Make sure that actual start and end time is set.
     *
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function setDates(array &$rows)
    {
        foreach ($rows as $index => &$row) {
            if (!isset($row[self::COLUMN_INPUT_ACTUAL_START])) {
                $row[self::COLUMN_INPUT_ACTUAL_START] = $row[self::COLUMN_INPUT_PLANNED_START];
            }
            if (!isset($row[self::COLUMN_INPUT_ACTUAL_END])) {
                $row[self::COLUMN_INPUT_ACTUAL_END] = $row[self::COLUMN_INPUT_PLANNED_END];
            }
        }
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculateOvertid(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculateNat(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculateIkkePlanlagt7(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculate50Pct(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculate13Timer(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculate11Timer(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculateDagenFør(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculate100Pct(array $row)
    {
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="",
     * )
     */
    private function calculateAntalHviletidsbrud(array $row)
    {
    }
}
