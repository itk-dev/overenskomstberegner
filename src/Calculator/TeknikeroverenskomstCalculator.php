<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Calculator;

use App\Calculator\Exception\InvalidArgumentException;
use DateTimeImmutable;
use DateTimeInterface;
use App\Annotation\Calculation;
use App\Annotation\Calculation\Placeholder;
use App\Annotation\Calculator;
use App\Annotation\Calculator\Argument;
use App\Annotation\Calculator\Setting;
use DateTime;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Note on date calculations: All date (and time) calculations in this class
 * are done using Excel timestamps, i.e. floating point numbers with the
 * integer part specifying the number of days since 1900-01-00 (yes, 00) and
 * the decimal part specifying the time of day. See
 * http://www.cpearson.com/excel/datetime.htm and similar resources for
 * details.
 *
 * @see http://excelformulabeautifier.com/
 *
 * @Calculator(
 *     name="Tekniker og lignende",
 *     description=""
 * )
 */
class TeknikeroverenskomstCalculator extends AbstractCalculator
{
    protected $name = 'Overenskomst for teknikere';

    /**
     * @Setting(type="string", name="Titel", description="Overskrift på resultat"),
     *
     * @var string
     */
    private $resultTitle;

    /**
     * @Setting(type="time", name="Overtid/nat fra", description="Tidspunkt hvor overtidsperioden starter"),
     *
     * @var DateTime
     */
    private $overtidNatFra;

    /**
     * @Setting(type="time", name="Overtid/nat til", description="Tidspunkt hvor overtidsperioden slutter"),
     *
     * @var DateTime
     */
    private $overtidNatTil;

    /**
     * @var float
     */
    private $_11_timer;

    /**
     * @var float
     */
    private $_13_timer;

    /**
     * @Setting(type="time", name="5571 start"),
     *
     * @var DateTime
     */
    private $_5571_start;

    /**
     * @Setting(type="time", name="5571 midt"),
     *
     * @var DateTime
     */
    private $_5571_midt;

    /**
     * @Setting(type="time", name="5571 slut"),
     *
     * @var DateTime
     */
    private $_5571_slut;

    /**
     * @Setting(type="time", name="6625 start"),
     *
     * @var DateTime
     */
    private $_6625_start;

    /**
     * @Setting(type="time", name="6625 slut"),
     *
     * @var DateTime
     */
    private $_6625_slut;

    /**
     * @Setting(type="time", name="Miljø start"),
     *
     * @var DateTime
     */
    private $miljoe_start;

    /**
     * @Setting(type="time", name="Miljø slut"),
     *
     * @var DateTime
     */
    private $miljoe_slut;

    /**
     * @Setting(type="time", name="Timeløn"),
     *
     * @var DateTime
     */
    private $timeloen;

    /**
     * @Setting(type="text", name="Normnedsættende events", description="Én per linje"),
     *
     * @var string
     */
    private $normnedsaettende;

    /**
     * @Setting(type="text", name="Kontraktnormer", description="CSV: kontrakt,ugenorm,normperiode"),
     *
     * @var string
     */
    private $kontraktnormer;

    /**
     * @Argument(type="date", name="Start time", description="Start time", required=true, default="first day of this month"),
     *
     * @var DateTime
     */
    private $startDate;

    /**
     * @Argument(type="date", name="End time", description="End time", required=true, default="last day of this month"),
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

    private const COLUMN_OUTPUT_P_5571 = 'P 5571';
    private const COLUMN_OUTPUT_P_6625 = 'P 6625';
    private const COLUMN_OUTPUT_P_MILJOE = 'P Miljø';
    private const COLUMN_OUTPUT_P_VARSEL = 'P Varsel';
    private const COLUMN_OUTPUT_P_DELT = 'P Delt';
    private const COLUMN_OUTPUT_P_50_PCT = 'P 50%';
    private const COLUMN_OUTPUT_P_100_PCT = 'P 100%';
    private const COLUMN_OUTPUT_P_ANTAL = 'P Antal';
    private const COLUMN_OUTPUT_P_NORMAL = 'P Normal';
    private const COLUMN_OUTPUT_TIMER2 = 'Timer2';
    private const COLUMN_OUTPUT_ARBEJDSDAGE = 'Arbejdsdage';

    private const COLUMN_SUM_P_5571 = 'Σ P 5571';
    private const COLUMN_SUM_P_6625 = 'Σ P 6625';
    private const COLUMN_SUM_P_MILJOE = 'Σ P Miljø';
    private const COLUMN_SUM_P_VARSEL = 'Σ P Varsel';
    private const COLUMN_SUM_P_DELT = 'Σ P Delt';
    private const COLUMN_SUM_P_50_PCT = 'Σ P 50%';
    private const COLUMN_SUM_P_100_PCT = 'Σ P 100%';
    private const COLUMN_SUM_P_ANTAL = 'Σ P Antal';
    private const COLUMN_SUM_P_NORMAL = 'Σ P Normal';
    private const COLUMN_SUM_P_TIMER2 = 'Σ Timer2';
    private const COLUMN_SUM_P_ARBEJDSDAGE = 'Σ Arbejdsdage';

    private const COLUMN_OUTPUT_EMPLOYEE_NUMBER = 'Medarbejdernummer';
    private const COLUMN_OUTPUT_LOENART = 'Lønart';
    private const COLUMN_OUTPUT_LOEBENR = 'Løbenr.';
    private const COLUMN_OUTPUT_ENHEDER_I_ALT = 'Enheder (i alt)';
    private const COLUMN_OUTPUT_IKRAFT_DATO = 'Ikraft dato (for lønmåned)';

    private const COLUMN_TEMP_OVERTID = 'overtid';
    private const COLUMN_TEMP_IS_OVERTIME = 'is overtime';
    private const COLUMN_TEMP_TIMER = 'timer';
    private const COLUMN_TEMP_NAT = 'nat';
    private const COLUMN_TEMP_IKKE_PLANLAGT7 = 'ikke planlagt 7';
    private const COLUMN_TEMP_13_TIMER = '13 timer';
    private const COLUMN_TEMP_11_TIMER = '11 timer';
    private const COLUMN_TEMP_5571 = '5571';
    private const COLUMN_TEMP_6625 = '6625';
    private const COLUMN_TEMP_MILJOE = 'miljø';
    private const COLUMN_TEMP_50_PCT = '50 %';
    private const COLUMN_TEMP_100_PCT = '100 %';
    private const COLUMN_TEMP_DAGEN_FOER = 'dagen før';
    private const COLUMN_TEMP_ANTAL_HVILETIDSBRUD = 'antal hviletidsbrud';

    private const COLUMN_TEST_REFERENCE_TIMER = 'test Timer';
    private const COLUMN_TEST_REFERENCE_OVERTID = 'test Overtid';
    private const COLUMN_TEST_REFERENCE_NAT = 'test Nat';
    private const COLUMN_TEST_REFERENCE_IKKE_PLANLAGT_7 = 'test Ikke pl./7';
    private const COLUMN_TEST_REFERENCE_50_PCT = 'test 50%';
    private const COLUMN_TEST_REFERENCE_13_TIMER = 'test 13 timer';
    private const COLUMN_TEST_REFERENCE_11_TIMER = 'test 11 timer';
    private const COLUMN_TEST_REFERENCE_DAGEN_FOER = 'test Dagen før';
    private const COLUMN_TEST_REFERENCE_100_PCT = 'test 100%';
    private const COLUMN_TEST_REFERENCE_ANTAL_HVILETIDSBRUD = 'test Antal Hv.';
    private const COLUMN_TEST_REFERENCE_5571 = 'test 5571';
    private const COLUMN_TEST_REFERENCE_6625 = 'test 6625';
    private const COLUMN_TEST_REFERENCE_MILJOE = 'test Miljø';
    private const COLUMN_TEST_REFERENCE_VARSEL = 'test Varsel';
    private const COLUMN_TEST_REFERENCE_DEL = 'test Delt';
    private const COLUMN_TEST_REFERENCE_HELLIGDAG = 'test Helligdag';
    private const COLUMN_TEST_REFERENCE_OT = 'test OT';
    private const COLUMN_TEST_REFERENCE_P_5571 = 'test P 5571';
    private const COLUMN_TEST_REFERENCE_P_6625 = 'test P 6625';
    private const COLUMN_TEST_REFERENCE_P_MILJOE = 'test P Miljø';
    private const COLUMN_TEST_REFERENCE_P_VARSEL = 'test P Varsel';
    private const COLUMN_TEST_REFERENCE_P_DELT = 'test P Delt';
    private const COLUMN_TEST_REFERENCE_P_50_PCT = 'test P 50%';
    private const COLUMN_TEST_REFERENCE_P_100_PCT = 'test P 100%';
    private const COLUMN_TEST_REFERENCE_P_ANTAL = 'test P Antal';
    private const COLUMN_TEST_REFERENCE_P_NORMAL = 'test P Normal';
    private const COLUMN_TEST_REFERENCE_TIMER2 = 'test Timer 2';
    private const COLUMN_TEST_REFERENCE_ARBEJDSDAGE = 'test Arbejdsdage';

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

        if ($this->testMode) {
            // TEST
            $dataColumnEnd = 'AL';
        }

        $dataRowStart = 3;
        $dataRowEnd = $sheet->getHighestRow('E');
        $dataRange = $dataColumnStart.$dataRowStart.':'.$dataColumnEnd.$dataRowEnd;

        $rows = $sheet->rangeToArray($dataRange, null, true, false);

        // DEBUG
        $rows = array_filter($rows, function ($index) {
            return true;

            return 13 === $index;
        }, ARRAY_FILTER_USE_KEY);
        // DEBUG

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
                    self::COLUMN_INPUT_DATE => $row[4],
                    self::COLUMN_INPUT_EVENT => $row[5],
                    self::COLUMN_INPUT_PLANNED_START => $row[6],
                    self::COLUMN_INPUT_PLANNED_END => $row[7],
                    self::COLUMN_INPUT_ACTUAL_START => $row[8],
                    self::COLUMN_INPUT_ACTUAL_END => $row[9],
                ];

                if ($this->testMode) {
                    $employeeRow += [
                        // TEST
                        self::COLUMN_TEST_REFERENCE_TIMER => $row[self::EXCEL_COLUMN_K],
                        self::COLUMN_TEST_REFERENCE_OVERTID => $row[self::EXCEL_COLUMN_L],
                        self::COLUMN_TEST_REFERENCE_NAT => $row[self::EXCEL_COLUMN_M],
                        self::COLUMN_TEST_REFERENCE_IKKE_PLANLAGT_7 => $row[self::EXCEL_COLUMN_N],
                        self::COLUMN_TEST_REFERENCE_50_PCT => $row[self::EXCEL_COLUMN_O],
                        self::COLUMN_TEST_REFERENCE_13_TIMER => $row[self::EXCEL_COLUMN_P],
                        self::COLUMN_TEST_REFERENCE_11_TIMER => $row[self::EXCEL_COLUMN_Q],
                        self::COLUMN_TEST_REFERENCE_DAGEN_FOER => $row[self::EXCEL_COLUMN_R],
                        self::COLUMN_TEST_REFERENCE_100_PCT => $row[self::EXCEL_COLUMN_S],
                        self::COLUMN_TEST_REFERENCE_ANTAL_HVILETIDSBRUD => $row[self::EXCEL_COLUMN_T],
                        self::COLUMN_TEST_REFERENCE_5571 => $row[self::EXCEL_COLUMN_U],
                        self::COLUMN_TEST_REFERENCE_6625 => $row[self::EXCEL_COLUMN_V],
                        self::COLUMN_TEST_REFERENCE_MILJOE => $row[self::EXCEL_COLUMN_W],
                        self::COLUMN_TEST_REFERENCE_VARSEL => $row[self::EXCEL_COLUMN_X],
                        self::COLUMN_TEST_REFERENCE_DEL => $row[self::EXCEL_COLUMN_Y],
                        self::COLUMN_TEST_REFERENCE_HELLIGDAG => $row[self::EXCEL_COLUMN_Z],
                        self::COLUMN_TEST_REFERENCE_OT => $row[self::EXCEL_COLUMN_AA],
                        self::COLUMN_TEST_REFERENCE_P_5571 => $row[self::EXCEL_COLUMN_AB],
                        self::COLUMN_TEST_REFERENCE_P_6625 => $row[self::EXCEL_COLUMN_AC],
                        self::COLUMN_TEST_REFERENCE_P_MILJOE => $row[self::EXCEL_COLUMN_AD],
                        self::COLUMN_TEST_REFERENCE_P_VARSEL => $row[self::EXCEL_COLUMN_AE],
                        self::COLUMN_TEST_REFERENCE_P_DELT => $row[self::EXCEL_COLUMN_AF],
                        self::COLUMN_TEST_REFERENCE_P_50_PCT => $row[self::EXCEL_COLUMN_AG],
                        self::COLUMN_TEST_REFERENCE_P_100_PCT => $row[self::EXCEL_COLUMN_AH],
                        self::COLUMN_TEST_REFERENCE_P_ANTAL => $row[self::EXCEL_COLUMN_AI],
                        self::COLUMN_TEST_REFERENCE_P_NORMAL => $row[self::EXCEL_COLUMN_AJ],
                        self::COLUMN_TEST_REFERENCE_TIMER2 => $row[self::EXCEL_COLUMN_AK],
                        self::COLUMN_TEST_REFERENCE_ARBEJDSDAGE => $row[self::EXCEL_COLUMN_AL],
                        // TEST
                    ];
                }

                // // Assume that actual end is on next day if less that actual start.
                // if ($employeeRow[self::COLUMN_INPUT_ACTUAL_END] < $employeeRow[self::COLUMN_INPUT_ACTUAL_START]) {
                //     $employeeRow[self::COLUMN_INPUT_ACTUAL_END] = $employeeRow[self::COLUMN_INPUT_ACTUAL_END]->add(new \DateInterval('P1D'));
                // }

                // Fill in missing dates.
                // if (!isset($employeeRow[self::COLUMN_INPUT_ACTUAL_START])) {
                //     $employeeRow[self::COLUMN_INPUT_ACTUAL_START] = $employeeRow[self::COLUMN_INPUT_PLANNED_START];
                // }
                // if (!isset($employeeRow[self::COLUMN_INPUT_ACTUAL_END])) {
                //     $employeeRow[self::COLUMN_INPUT_ACTUAL_END] = $employeeRow[self::COLUMN_INPUT_PLANNED_END];
                // }

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

        if ($this->endDate <= $this->startDate) {
            throw new InvalidArgumentException('End date must be after start date');
        }

        // Convert times to Excel floats.
        $this->overtidNatFra = $this->time2Excel($this->overtidNatFra);
        $this->overtidNatTil = $this->time2Excel($this->overtidNatTil);
        $this->_5571_start = $this->time2Excel($this->_5571_start);
        $this->_5571_midt = $this->time2Excel($this->_5571_midt);
        $this->_5571_slut = $this->time2Excel($this->_5571_slut);
        $this->_6625_start = $this->time2Excel($this->_6625_start);
        $this->_6625_slut = $this->time2Excel($this->_6625_slut);
        $this->miljoe_start = $this->time2Excel($this->miljoe_start);
        $this->miljoe_slut = $this->time2Excel($this->miljoe_slut);
        $this->_11_timer = $this->time2Excel(new DateTime('@0 11:00'));
        $this->_13_timer = $this->time2Excel(new DateTime('@0 13:00'));
        $this->timeloen = $this->time2Excel($this->timeloen);

        $startDate = $this->dateTime2Excel($this->startDate);
        $endDate = $this->dateTime2Excel($this->endDate) + 1;

        foreach ($this->data as $employeeNumber => &$rows) {
            $result = $this->calculateEmployee($rows);

            // Keep only rows in the specified report date interval.
            $result = array_values(array_filter($result, function (array $row) use ($startDate, $endDate) {
                return $startDate <= $row[self::COLUMN_INPUT_DATE]
                    && $row[self::COLUMN_INPUT_DATE] < $endDate;
            }));

            if ($this->testMode) {
                $this->testCheckRows($rows);
            }

            $row = $rows[0];

            // Compute non-zero sums.
            $result = array_filter([
                self::COLUMN_SUM_P_5571 => array_sum(array_column($result, self::COLUMN_OUTPUT_P_5571)),
                self::COLUMN_SUM_P_6625 => array_sum(array_column($result, self::COLUMN_OUTPUT_P_6625)),
                self::COLUMN_SUM_P_MILJOE => array_sum(array_column($result, self::COLUMN_OUTPUT_P_MILJOE)),
                self::COLUMN_SUM_P_VARSEL => array_sum(array_column($result, self::COLUMN_OUTPUT_P_VARSEL)),
                self::COLUMN_SUM_P_DELT => array_sum(array_column($result, self::COLUMN_OUTPUT_P_DELT)),
                self::COLUMN_SUM_P_50_PCT => array_sum(array_column($result, self::COLUMN_OUTPUT_P_50_PCT)),
                self::COLUMN_SUM_P_100_PCT => array_sum(array_column($result, self::COLUMN_OUTPUT_P_100_PCT)),
                self::COLUMN_SUM_P_ANTAL => array_sum(array_column($result, self::COLUMN_OUTPUT_P_ANTAL)),
                self::COLUMN_SUM_P_NORMAL => array_sum(array_column($result, self::COLUMN_OUTPUT_P_NORMAL)),
                self::COLUMN_SUM_P_TIMER2 => array_sum(array_column($result, self::COLUMN_OUTPUT_TIMER2)),
                self::COLUMN_SUM_P_ARBEJDSDAGE => array_sum(array_column($result, self::COLUMN_OUTPUT_ARBEJDSDAGE)),
            ]);

            if (!empty($result)) {
                $result = [
                    self::COLUMN_INPUT_NAME => $row[self::COLUMN_INPUT_NAME],
                    self::COLUMN_INPUT_EMPLOYEE_NUMBER => $row[self::COLUMN_INPUT_EMPLOYEE_NUMBER],
                ] + $result;

                $this->result[$employeeNumber] = $result;
            }
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

        // Sum af P 5571
        // Sum af P 6625
        // Sum af P Miljø
        // Sum af P Varsel
        // Sum af P Delt
        // Sum af P 50%
        // Sum af P 100%
        // Sum af P Antal
        // Sum af P Normal

        $result = new Spreadsheet();
        $sheet = $result->getActiveSheet();
        $rowIndex = 1;
        $this->writeCell($sheet, 1, 1, $this->resultTitle, 5);
        ++$rowIndex;

        foreach ($this->result as $employeeNumber => $row) {
            $columnIndex = 1;
            foreach ($row as $key => $value) {
                $sheet->setCellValueByColumnAndRow($columnIndex, $rowIndex, $key);
                $sheet->setCellValueByColumnAndRow($columnIndex + 1, $rowIndex, $value);
                $columnIndex += 2;
            }
            ++$rowIndex;
        }

        return $result;

        $this->writeCells($sheet, 1, $rowIndex, [
            self::COLUMN_OUTPUT_EMPLOYEE_NUMBER,
            self::COLUMN_OUTPUT_LOENART,
            self::COLUMN_OUTPUT_LOEBENR,
            self::COLUMN_OUTPUT_ENHEDER_I_ALT,
            self::COLUMN_OUTPUT_IKRAFT_DATO,
        ]);
        ++$rowIndex;

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

    private function calculateEmployee(array &$rows)
    {
        $this->setRows($rows);
        // Some values must be calculated before other stuff.
        while ($this->nextRow()) {
            $this->calculate13Timer();
            $this->calculateDagenFoer();
        }

        $this->setRows($rows);
        while ($this->nextRow()) {
            $this->calculateP5571();
            $this->calculateP6625();
            $this->calculatePMiljoe();
            $this->calculatePVarsel();
            $this->calculatePDelt();
            $this->calculateP50Pct();
            $this->calculateP100Pct();
            $this->calculatePAntal();
            $this->calculatePNormal();
        }

        return $rows;

        // return [
        //     self::COLUMN_OUTPUT_EMPLOYEE_NUMBER => $employeeNumber,
        //     self::COLUMN_OUTPUT_LOENART => null,
        //     self::COLUMN_OUTPUT_LOEBENR => null,
        //     self::COLUMN_OUTPUT_ENHEDER_I_ALT => null,
        //     self::COLUMN_OUTPUT_IKRAFT_DATO => null,
        // ];
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
     =IF(
     ISERROR(
     VLOOKUP(
     E3,
     Helligdage!$B:$B,
     1,
     FALSE
     )
     ),
     IF(
     F3 = ""Vagt"",
     IF(
     WEEKDAY(
     E3,
     2
     ) < 6,
     IF(
     I3 < E3 + Meta!$D$4,
     IF(
     J3 < E3 + Meta!$D$4,
     J3 - I3,
     IF(
     J3 < E3 + Meta!$D$3,
     E3 + Meta!$D$4 - I3,
     E3 + Meta!$D$4 - I3 + J3 - E3 + Meta!$D$3
     )
     ),
     IF(
     J3 < E3 + Meta!$D$3,
     0,
     IF(
     J3 < E3 + 1 + Meta!$D$4,
     J3 - E3 - Meta!$D$3,
     1 + Meta!$D$4 - I3
     )
     )
     ),
     IF(
     WEEKDAY(
     E3,
     2
     ) = 6,
     IF(
     I3 < E3 + Meta!$D$4,
     IF(
     J3 < E3 + Meta!$D$4,
     J3 - I3,
     0
     ),
     IF(
     J3 < E3 + Meta!$D$2,
     J3 - I3,
     IF(
     I3 < E3 + Meta!$D$2,
     IF(
     J3 < E3 + Meta!$D$3,
     J3 - ( E3 + Meta!$D$2 ),
     E3 + Meta!$D$3 - ( E3 + Meta!$D$2 )
     ),
     IF(
     J3 < E3 + Meta!$D$3,
     J3 - I3,
     E3 + Meta!$D$3 - I3
     )
     )
     )
     ),
     0
     )
     ),
     0
     ),
     0
     )
     "
     * )
     */
    private function calculate5571()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_5571, function () {
            if ($this->isHoliday($this->get(self::COLUMN_INPUT_DATE))) {
                return 0;
            }

            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $_5571_start = $this->_5571_start;
            $_5571_midt = $this->_5571_midt;
            $_5571_slut = $this->_5571_slut;
            $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START); // I
            $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END); // J

            if (self::EVENT_VAGT === $event) {
                if ($this->getWeekday($date) < self::WEEKDAY_SATURDAY) {
                    if ($actual_start < $date + $_5571_slut) {
                        if ($actual_end < $date + $_5571_slut) {
                            $actual_end - $actual_start;
                        } else {
                            if ($actual_end < $date + $_5571_midt) {
                                return $date + $_5571_slut - $actual_start;
                            } else {
                                return $date + $_5571_slut - $actual_start + $actual_end - $date + $_5571_midt;
                            }
                        }
                    } else {
                        if ($actual_end < $date + $_5571_midt) {
                            return 0;
                        } else {
                            if ($actual_end < $date + 1 + $_5571_slut) {
                                return $actual_end - $date - $_5571_midt;
                            } else {
                                return 1 + $_5571_slut - $actual_start;
                            }
                        }
                    }
                } else {
                    if (self::WEEKDAY_SATURDAY === $this->getWeekday($date)) {
                        if ($actual_start < $date + $_5571_slut) {
                            if ($actual_end < $date + $_5571_slut) {
                                return $actual_end - $actual_start;
                            } else {
                                return 0;
                            }
                        } else {
                            if ($actual_end < $date + $_5571_start) {
                                return $actual_end - $actual_start;
                            } else {
                                if ($actual_start < $date + $_5571_start) {
                                    if ($actual_end < $date + $_5571_midt) {
                                        return $actual_end - ($date + $_5571_start);
                                    } else {
                                        return $date + $_5571_midt - ($date + $_5571_start);
                                    }
                                } else {
                                    if ($actual_end < $date + $_5571_midt) {
                                        return $actual_end - $actual_start;
                                    } else {
                                        return $date + $_5571_midt - $actual_start;
                                    }
                                }
                            }
                        }
                    } else {
                        return 0;
                    }
                }
            } else {
                return 0;
            }
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="=IF(AND(Startdato<=Data!$E3;Data!$E3<=Slutdato);Data!U3*24;"")"
     * )
     */
    private function calculateP5571()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_5571, function () {
            return 24 * $this->calculate5571();
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
     =IF(
     Z3 = ""X"",
     IF(
     OR(
     J3 < E3 + 1,
     WEEKDAY(
     E3 + 1,
     2
     ) = 7,
     OFFSET(
     Z3,
     1,
     0
     ) = ""X""
     ),
     J3 - I3,
     IF(
     WEEKDAY(
     E3 + 1,
     2
     ) = 1,
     IF(
     J3 < E3 + 1 + Meta!$E$3,
     J3 - I3,
     E3 + 1 + Meta!$E$3
     ),
     0
     )
     ),
     IF(
     F3 = ""Vagt"",
     IF(
     OR(
     WEEKDAY(
     E3,
     2
     ) >= 6,
     WEEKDAY(
     E3,
     2
     ) = 1
     ),
     IF(
     WEEKDAY(
     E3,
     2
     ) = 6,
     IF(
     I3 < E3 + Meta!$E$2,
     IF(
     J3 <= E3 + Meta!$E$2,
     0,
     J3 - ( E3 + Meta!$E$2 )
     ),
     J3 - I3
     ),
     IF(
     WEEKDAY(
     E3,
     2
     ) = 7,
     IF(
     J3 < E3 + 1 + Meta!$E$3,
     J3 - I3,
     E3 + 1 + Meta!$E$3 - I3
     ),
     IF(
     I3 > E3 + Meta!$E$3,
     0,
     IF(
     J3 > E3 + Meta!$E$3,
     E3 + Meta!$E$3 - I3,
     J3 - I3
     )
     )
     )
     ),
     0
     ),
     0
     )
     )
     "
     * )
     */
    private function calculate6625()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_6625, function () {
            $date = $this->get(self::COLUMN_INPUT_DATE); // E
            $date_prev = $this->get(self::COLUMN_INPUT_DATE, -1);
            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START); // I
            $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END); // J
            $_6625_start = $this->_6625_start; // Meta!$E$2
            $_6625_slut = $this->_6625_slut; // Meta!$E$3

            if ($this->isHoliday($date)) {
                if ($actual_end < $date + 1
                    || self::WEEKDAY_SUNDAY === $this->getWeekday($date + 1)
                    || $this->isHoliday($date_prev)) {
                    return $actual_end - $actual_start;
                } else {
                    if (self::WEEKDAY_MONDAY === $this->getWeekday($date + 1)) {
                        if ($actual_end < $date + 1 + $_6625_slut) {
                            return $actual_end - $actual_start;
                        } else {
                            return $date + 1 + $_6625_slut;
                        }
                    } else {
                        return 0;
                    }
                }
            } else {
                if (self::EVENT_VAGT === $event) {
                    if ($this->getWeekday($date) >= self::WEEKDAY_SATURDAY
                        || self::WEEKDAY_MONDAY === $this->getWeekday($date)) {
                        if (self::WEEKDAY_SATURDAY === $this->getWeekday($date)) {
                            if ($actual_start < $date + $_6625_start) {
                                if ($actual_end <= $date + $_6625_start) {
                                    return                                    0;
                                } else {
                                    return $actual_end - ($date + $_6625_start);
                                }
                            } else {
                                return $actual_end - $actual_start;
                            }
                        } else {
                            if (self::WEEKDAY_SUNDAY === $this->getWeekday($date)) {
                                if ($actual_end < $date + 1 + $_6625_slut) {
                                    return $actual_end - $actual_start;
                                } else {
                                    return $date + 1 + $_6625_slut - $actual_start;
                                }
                            } else {
                                if (
                                    $actual_start > $date + $_6625_slut) {
                                    return 0;
                                } else {
                                    if ($actual_end > $date + $_6625_slut) {
                                        return $date + $_6625_slut - $actual_start;
                                    } else {
                                        return $actual_end - $actual_start;
                                    }
                                }
                            }
                        }

                        return 0;
                    }

                    return 0;
                }
            }
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculateP6625()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_6625, function () {
            return 24 * $this->calculate6625();
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="=IF(F3=""Vagt"",IF(I3<E3+Meta!$F$3,IF(J3<E3+Meta!$F$3,J3-I3,IF(J3<=E3+MiljoStart,E3+Meta!$F$3-I3,(E3+Meta!$F$3-I3)+(J3-(E3+MiljoStart)))),IF(I3<E3+MiljoStart,IF(J3<=E3+MiljoStart,0,J3-(E3+MiljoStart)),IF(J3<=E3+Meta!$F$3+1,J3-I3,E3+Meta!$F$3+1-I3))),0)*8.1%"
     * )
     */
    private function calculateMiljoe()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_MILJOE, function () {
            $calculate = function () {
                $date = $this->get(self::COLUMN_INPUT_DATE); // E
                $event = $this->get(self::COLUMN_INPUT_EVENT); // F
                $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START); // I
                $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END); // J
                $miljoe_start = $this->miljoe_start; // Meta!$F$2
                $miljoe_slut = $this->miljoe_slut; // Meta!$F$3

                if (self::EVENT_VAGT === $event) {
                    if ($actual_start < $date + $miljoe_slut) {
                        if ($actual_end < $date + $miljoe_slut) {
                            return $actual_end - $actual_start;
                        } else {
                            if ($actual_end <= $date + $miljoe_start) {
                                return $date + $miljoe_slut - $actual_start;
                            } else {
                                return ($date + $miljoe_slut - $actual_start) + ($actual_end - ($date + $miljoe_start));
                            }
                        }
                    } else {
                        if ($actual_start < $date + $miljoe_start) {
                            if ($actual_end <= $date + $miljoe_start) {
                                return 0;
                            } else {
                                return $actual_end - ($date + $miljoe_start);
                            }
                        } else {
                            if ($actual_end <= $date + $miljoe_slut + 1) {
                                return $actual_end - $actual_start;
                            } else {
                                return $date + $miljoe_slut + 1 - $actual_start;
                            }
                        }
                    }
                } else {
                    return 0;
                }
            };

            return $calculate() * 0.081;
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculatePMiljoe()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_MILJOE, function () {
            return 24 * $this->calculateMiljoe();
        });
    }

//    /**
//     * @Calculation(
//     *     name="arbejdstimer",
//     *     description="",
//     *     formula="",
//     *     overenskomsttekst="",
//     *     excelFormula=""
//     * )
//     */
//    private function calculateVarsel()
//    {
//        return $this->calculateColumn(self::COLUMN_TEMP_VARSEL, function () {
//            return $this->calculateVarsel();
//        });
//    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculatePVarsel()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_VARSEL, function () {
            return self::EVENT_LOEN_VARSEL === $this->get(self::COLUMN_INPUT_EVENT) ? 1 : null;
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculatePDelt()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_DELT, function () {
            return self::EVENT_LOEN_DELT_TJENESTE === $this->get(self::COLUMN_INPUT_EVENT) ? 1 : null;
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
=IF(
    AND(
        HLOOKUP(
            VLOOKUP(
                Data!C3,
                Overenskomst,
                3,
                FALSE
            ),
            Periode,
            2,
            FALSE
        ) <= Data!$E3,
        Data!$E3 <=
        HLOOKUP(
            VLOOKUP(
                Data!C3,
                Overenskomst,
                3,
                FALSE
            ),
            Periode,
            3,
            FALSE
        )
    ),
    O3 * 24,
    """"
)
"
     * )
     */
    private function calculateP50Pct()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_50_PCT, function () {
            $contract = $this->get(self::COLUMN_INPUT_CONTRACT);
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $normperiode = $this->getNormperiode($contract);

            if ($normperiode[0] <= $date && $date <= $normperiode[1]) {
                return $this->calculate50Pct() * 24;
            } else {
                return null;
            }
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculateP100Pct()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_100_PCT, function () {
            $contract = $this->get(self::COLUMN_INPUT_CONTRACT);
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $normperiode = $this->getNormperiode($contract);

            if ($normperiode[0] <= $date && $date <= $normperiode[1]) {
                return $this->calculate100Pct() * 24;
            } else {
                return null;
            }
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculatePAntal()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_ANTAL, function () {
            // @TODO Are `24 * ` missing here?
            return $this->calculateAntalHviletidsbrud();
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
=IF(
    AND(
        HLOOKUP(
            VLOOKUP(
                Data!C3,
                Overenskomst,
                3,
                FALSE
            ),
            Periode,
            2,
            FALSE
        ) <= Data!$E3,
        Data!$E3 <=
        HLOOKUP(
            VLOOKUP(
                Data!C3,
                Overenskomst,
                3,
                FALSE
            ),
            Periode,
            3,
            FALSE
        )
    ),
    IF(
        K3 > 0,
        K3 * 24 -
        SUM(
            AG3:AH3
        ),
        0
    ),
    """"
)
"
     * )
     */
    private function calculatePNormal()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_NORMAL, function () {
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $contract = $this->get(self::COLUMN_INPUT_CONTRACT);
            $normperiode = $this->getNormperiode($contract);

            if ($normperiode[0] <= $date && $date <= $normperiode[1]) {
                $timer = $this->calculateTimer();
                $p50Pct = $this->calculateP50Pct();
                $p100Pct = $this->calculateP100Pct();
                if ($timer > 0) {
                    return $timer * 24 - ($p50Pct + $p100Pct);
                } else {
                    return   0;
                }
            } else {
                return null;
            }
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculateTimer2()
    {
        // @TODO: Is this actually used?
        return $this->calculateColumn(self::COLUMN_OUTPUT_TIMER2, function () {
            throw new \RuntimeException(__METHOD__.' not implemented');
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     */
    private function calculateArbejdsdage()
    {
        // @TODO: Is this actually used?
        return $this->calculateColumn(self::COLUMN_OUTPUT_ARBEJDSDAGE, function () {
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $datePrev = $this->getPrev(self::COLUMN_INPUT_DATE);

            // @TODO: This does not look right!
            return $date !== $datePrev ? 1 : 0;
        });
    }

    // @TODO: Is this the complete list?
    private const EVENT_VAGT = 'Vagt';
    private const EVENT_SYGDOM = 'Sygdom';
    private const EVENT_LOEN_OVERTID = 'Løn: Overtid';
    private const EVENT_LOEN_IKKE_PLANLAGT_7_DAG = 'Løn: Ikke planlagt/7. dag';
    private const EVENT_KURSUS = 'Kursus';
    private const EVENT_FERIETIMER = 'Ferietimer';
    private const EVENT_SENIORDAG = 'Seniordag';
    private const EVENT_LOEN_DELT_TJENESTE = 'Løn: Delt tjeneste';
    // @TODO Is this event actually used?
    private const EVENT_LOEN_VARSEL = 'Løn: Varsel';

    // @TODO: Is this the complete list?
    // private const CONTRACT_TEKNIK_37_HOURS_3_MÅNEDER = 'Teknik 37 hours 3 måneder';
    private const CONTRACT_TEKNIK_37_HOURS = 'Teknik 37 hours';
    private const CONTRACT_TEKNIK_32_HOURS = 'Teknik 32 hours';
    private const CONTRACT_TIMELOENNEDE = 'Timelønnede';

    /**
     * Calculate arbejdstimer.
     *
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="\frac{\text{overarbejdeFra}}{\text{overarbejdeTil}}",
     *     placeholders={
     *         "a": @Placeholder(name="The a value", description="Value of a", type="int"),
     *         "b": @Placeholder(name="The a value", description="Value of a", type="int"),
     *     },
     *     overenskomsttekst="…",
     *     excelFormula="=
     HVIS(
     ELLER(
     F3 = ""Vagt"";
     OG(
     F3 = ""Sygdom"";
     J3 <> 0
     )
     );
     J3 - I3;
     HVIS(
     ER.FEJL(
     LOPSLAG(
     F3;
     Normnedsættende;
     1;
     FALSK
     )
     );
     0;
     LOPSLAG(
     C3;
     Meta!H:I;
     2;
     FALSK
     ) / 5 / 24
     )
     )
     ",
     * )
     */
    private function calculateTimer()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_TIMER, function () {
            $contract = $this->get(self::COLUMN_INPUT_CONTRACT);
            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START);
            $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END);

            if (self::EVENT_VAGT === $event
                || (self::EVENT_SYGDOM === $event && !empty($actual_end))) {
                return $actual_end - $actual_start;
            } elseif (!$this->isNormnedsaettende($event)) {
                return 0;
            } else {
                return $this->getUgenorm($contract) / 5 / 24;
            }
        });
    }

    private $normnedsaettendeItems;

    private function isNormnedsaettende($event)
    {
        if (null === $this->normnedsaettendeItems) {
            $this->normnedsaettendeItems = array_filter(array_map('trim', explode(PHP_EOL, $this->normnedsaettende)));
        }

        return \in_array($event, $this->normnedsaettendeItems);
    }

    private $kontraktnormerItems;

    private function kontraktnormer(string $contract = null)
    {
        if (null === $this->kontraktnormerItems) {
            $this->kontraktnormerItems = array_column(array_map('str_getcsv', array_filter(array_map('trim', explode(PHP_EOL, $this->kontraktnormer)))), null, 0);
        }

        if (null !== $contract) {
            if (!isset($this->kontraktnormerItems[$contract])) {
                throw new \RuntimeException(sprintf('Invalid contract: %s', $contract));
            }

            return $this->kontraktnormerItems[$contract];
        }

        return $this->kontraktnormerItems;
    }

    private function getUgenorm($contract)
    {
        $norm = $this->kontraktnormer($contract);

        return (int) $norm[1];
    }

    /**
     * @return \DateTimeInterface[]
     */
    private function getNormperiode($contract, $asExcelDates = true)
    {
        $calculate = function ($period) {
            switch ($period) {
                case 1:
                    $offset = $this->startDate->format(DateTimeInterface::ATOM);

                    return [
                        new DateTimeImmutable($offset.' first day of month'),
                        new DateTimeImmutable($offset.' last day of month'),
                    ];
                case 3:
                    // Get quarter containing start date.
                    $month = (int) $this->startDate->format('n');
                    $startQuarterMonth = 3 * (int) floor(($month - 1) / 3) + 1;

                    return [
                        new DateTimeImmutable($this->startDate->format(sprintf('Y-%02d-d\TH:i:sP', $startQuarterMonth)).' first day of month'),
                        new DateTimeImmutable($this->startDate->format(sprintf('Y-%02d-d\TH:i:sP', $startQuarterMonth + 2)).' last day of month'),
                    ];

                default:
                    throw new \RuntimeException(sprintf('Invalid norm period: %d', $period));
            }
        };

        // @TODO: What to do if contract is not set?
        $norm = $this->kontraktnormer($contract ?? '');
        $result = $calculate((int) $norm[2]);

        if ($asExcelDates) {
            $result = array_map([$this, 'dateTime2Excel'], $result);
        }

        return $result;
    }

    /**
     * Calculate if a row is overtime.
     *
     * @Calculation(
     *     name="Mellemregning: OT (overtid)",
     *     description="Formlen kigge 2 celler op og ned for at se, om datoen er den samme og om $event = ”Løn: Overtid” og returnerer ”OT”, hvis dette er tilfældet.",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     OG(
     E7 = E6;
     F6 = ""Løn: Overtid""
     );
     ""OT"";
     HVIS(
     OG(
     E7 = E5;
     F5 = ""Løn: Overtid""
     );
     ""OT"";
     HVIS(
     OG(
     E7 = E8;
     F8 = ""Løn: Overtid""
     );
     ""OT"";
     HVIS(
     OG(
     E7 = E9;
     F9 = ""Løn: Overtid""
     );
     ""OT"";
     """"
     )
     )
     )
     )
     ",
     * )
     */
    private function calculateIsOvertime()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_IS_OVERTIME, function () {
            return ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -2) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, -2))
                || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -1) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, -1))
                || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +1) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, +1))
                || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +2) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, +2));
        });
    }

    /**
     * @Calculation(
     *     name="Overtid",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="§ 6. Overarbejde/deltidsbeskæftigedes merarbejde
     Stk. 1
     Arbejde ud over den ved tjenestelisten fastlagte daglige arbejdstid for en fuldtidsbeskæftiget betragtes som overarbejde.
     Stk. 2
     Arbejde mellem kl. 00.00 - 08.00 betragtes altid som overarbejde. [Se Nat-beregning]
     Stk. 3
     For deltidsansatte er arbejde udover 8 timer dagligt overarbejde, dog ikke hvis tjenesten er planlagt til over 8 timer.
     Stk. 4
     Overarbejde søges godtgjort med frihed (afspadsering), der skal være af samme varighed, som det præsterede overarbejde med tillæg af 50% op rundet til antal hele timer",
     *     excelFormula="=
     HVIS(
     C3 <> ""Timelønnede"";
     HVIS(
     OG(
     J3 > H3;
     N3 = 0;
     AA3 = ""OT""
     );
     HVIS(
     J3 < E3 + 1;
     J3 - H3;
     HVIS(
     H3 >= E3 + 1;
     0;
     E3 + Meta!$A$2 - H3
     )
     );
     0
     );
     HVIS(
     OG(
     J3 > H3;
     K3 > Meta!$G$2
     );
     HVIS(
     J3 < E3 + 1;
     HVIS(
     H3 - G3 < Meta!$G$2;
     K3 - Meta!$G$2;
     K3 - ( H3 - G3 )
     );
     HVIS(
     J3 < E3 + 1 + Meta!$A$3;
     HVIS(
     H3 - G3 < Meta!$G$2;
     J3 - Meta!$G$2 - M3;
     K3 - ( H3 - G3 ) - M3
     );
     HVIS(
     I3 < E3 + 21 / 24;
     0;
     J3 - ( E3 + 1 + Meta!$A$3 )
     )
     )
     );
     0
     )
     )
     ",
     * )
     */
    private function calculateOvertid()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_OVERTID, function () {
            $contract = $this->get(self::COLUMN_INPUT_CONTRACT);
            $planned_start = $this->get(self::COLUMN_INPUT_PLANNED_START);
            $planned_end = $this->get(self::COLUMN_INPUT_PLANNED_END);
            $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END);
            $OT = $this->calculateIsOvertime();
            $timeloen = $this->timeloen;
            $timer = $this->calculateTimer();
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $overtid_start = $this->overtidNatFra;
            $overtid_slut = $this->overtidNatTil;
            $nat = (int) $this->calculateNat();

            if (self::CONTRACT_TIMELOENNEDE !== $contract) {
                if ($actual_end > $planned_end && $OT) {
                    if ($actual_end < $date + 1) {
                        return $actual_end - $planned_end;
                    } else {
                        if ($planned_end >= $date + 1) {
                            return 0;
                        } else {
                            return $date + $overtid_slut - $planned_end;
                        }

                        return 0;
                    }
                } else {
                    if ($actual_end > $planned_end && $timer > $timeloen) {
                        if ($actual_end < $date + 1) {
                            if ($planned_end - $planned_start < $timeloen) {
                                return $timer - $timeloen;
                            } else {
                                return $timer - $planned_end - $planned_start;
                            }
                        } else {
                            if ($actual_end < $date + 1 + $overtid_start) {
                                if ($planned_end - $planned_start < $overtid_slut) {
                                    return $actual_end - $overtid_slut - $nat;
                                } else {
                                    return $timer - $planned_end - $planned_start - $nat;
                                }
                                if ($actual_start < $date + 21 / 24) {
                                    return 0;
                                } else {
                                    return $actual_end - ($date + 1 + $overtid_start);
                                }
                            }
                        }

                        return 0;
                    }
                }
            }
        });
    }

    /**
     * @Calculation(
     *     name="Nat",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="Stk. 2
     Arbejde mellem kl. 00.00 - 08.00 betragtes altid som overarbejde. [Se Nat-beregning]
     Stk. 4
     Overarbejde søges godtgjort med frihed (afspadsering), der skal være af samme varighed, som det præsterede overarbejde med tillæg af 50% op rundet til antal hele timer",
     *     excelFormula="=
     HVIS(
     OG(
     F7 = ""Vagt"";
     N7 = 0
     );
     HVIS(
     I7 < E7 + Meta!$A$3;
     HVIS(
     J7 < E7 + Meta!$A$3;
     J7 - I7;
     E7 + Meta!$A$3 - I7
     );
     HVIS(
     J7 < E7 + 1;
     0;
     HVIS(
     J7 < E7 + 1 + Meta!$A$3;
     J7 - ( E7 + 1 );
     Meta!$A$3
     )
     )
     );
     0
     )
     ",
     * )
     */
    private function calculateNat()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_NAT, function () {
            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $ikkePlanlagt7 = $this->calculateIkkePlanlagt7();
            $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START);
            $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END);
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $overtid_start = $this->overtidNatFra;

            if (self::EVENT_VAGT === $event && 0 === $ikkePlanlagt7) {
                if ($actual_start < $date + $overtid_start) {
                    if ($actual_end < $date + $overtid_start) {
                        return $actual_end - $actual_start;
                    } else {
                        return $date + $overtid_start - $actual_start;
                    }
                } else {
                    if ($actual_end < $date + 1) {
                        return 0;
                    } else {
                        if ($actual_end < $date + 1 + $overtid_start) {
                            return $actual_end - ($date + 1);
                        } else {
                            return $overtid_start;
                        }
                    }
                }
            } else {
                return 0;
            }
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     OG(E3 = E2; F2 = ""Løn: Ikke planlagt/7. dag"");
     J3 - I3;
     HVIS(
     OG(E3 = E1; F1 = ""Løn: Ikke planlagt/7. dag"");
     J3 - I3;
     HVIS(
     OG(E3 = E4; F4 = ""Løn: Ikke planlagt/7. dag"");
     J3 - I3;
     HVIS(
     OG(E3 = E5; F5 = ""Løn: Ikke planlagt/7. dag"");
     J3 - I3;
     0
     )
     )
     )
     )
     ",
     * )
     */
    private function calculateIkkePlanlagt7()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_IKKE_PLANLAGT7, function () {
            return (($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -2) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, -2))
                    || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -1) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, -1))
                    || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +1) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, +1))
                    || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +2) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, +2)))
                ? $this->get(self::COLUMN_INPUT_ACTUAL_END) - $this->get(self::COLUMN_INPUT_ACTUAL_END) : 0;
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     OG(
     F3 = ""Vagt"";
     SUM(
     L3:N3
     ) > 0
     );
     HVIS(
     S3 = 0;
     SUM(
     L3:N3
     );
     HVIS(
     S3 >=
     SUM(
     L3:N3
     );
     0;
     SUM(
     L3:N3
     ) - S3
     )
     );
     0
     )
     ",
     * )
     */
    private function calculate50Pct()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_50_PCT, function () {
            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $sum = $this->calculateOvertid() + $this->calculateNat() + $this->calculateIkkePlanlagt7();
            $_100Pct = $this->calculate100Pct();

            if (self::EVENT_VAGT === $event && $sum > 0) {
                if (0 === $_100Pct) {
                    return $sum;
                } else {
                    if ($_100Pct >= $sum) {
                        return 0;
                    } else {
                        return $sum - $_100Pct;
                    }
                }
            } else {
                return 0;
            }
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     OG(
     F3 = ""Vagt"";
     K3 <> 0;
     K3 > Meta!$B$2 + 0,01
     );
     K3 - Meta!$B$2;
     0
     )
     ",
     * )
     */
    private function calculate13Timer()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_13_TIMER, function () {
            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $timer = $this->calculateTimer(); // K
            $_13_timer = $this->_13_timer;

            if (self::EVENT_VAGT === $event and 0 !== $timer and $timer > $_13_timer) {
                return $timer - $_13_timer;
            } else {
                return 0;
            }
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     ELLER(
     F3 <> ""Vagt"";
     FORSKYDNING(
     E3;
     -1;
     0
     ) <> E3 - 1
     );
     0;
     HVIS(
     FORSKYDNING(
     F3;
     -1;
     0
     ) <> ""Vagt"";
     HVIS (            OG(
     FORSKYDNING(
     E3;
     -2;
     0
     ) <> E3 - 1;
     FORSKYDNING(
     F3;
     -2;
     0
     ) <> ""Vagt""
     ) ; 0 ;
     HVIS(
     I3 - J1 < _11_timer;
     _11_timer - ( Data!I3 - Data!J1 );
     0
     ) );
     HVIS(
     I3 - J2 < _11_timer;
     _11_timer - ( Data!I3 - Data!J2 );
     0
     )
     )
     )
     ",
     * )
     */
    private function calculate11Timer()
    {
        $calculate = function () {
            $event = $this->get(self::COLUMN_INPUT_EVENT);
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $date_prev = $this->get(self::COLUMN_INPUT_DATE, -1);
            $date_prev_prev = $this->get(self::COLUMN_INPUT_DATE, -2);
            $event_prev = $this->get(self::COLUMN_INPUT_EVENT, -1);
            $event_prev_prev = $this->get(self::COLUMN_INPUT_EVENT, -2);
            $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START);
            $actual_end_prev = $this->get(self::COLUMN_INPUT_ACTUAL_END, -1);
            $actual_end_prev_prev = $this->get(self::COLUMN_INPUT_ACTUAL_END, -2);
            $_11_timer = $this->_11_timer;

            if (self::EVENT_VAGT !== $event || $date_prev !== $date - 1) {
                return 0;
            } else {
                if (self::EVENT_VAGT !== $event_prev) {
                    if ($date_prev_prev !== $date - 1 && self::EVENT_VAGT !== $event_prev_prev) {
                        return 0;
                    } else {
                        if ($actual_start - $actual_end_prev_prev < $_11_timer) {
                            return $_11_timer - ($actual_start - $actual_end_prev_prev);
                        } else {
                            return 0;
                        }
                    }
                } else {
                    if ($actual_start - $actual_end_prev < $_11_timer) {
                        return $_11_timer - ($actual_start - $actual_end_prev);
                    } else {
                        return 0;
                    }
                }
            }
        };

        $this->set(self::COLUMN_TEMP_11_TIMER, $calculate());
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     OG(
     E3 - 1 =
     FORSKYDNING(
     E3;
     -1;
     0
     );
     P2 > 0
     );
     P2;
     HVIS(
     OG(
     E3 - 1 =
     FORSKYDNING(
     E3;
     -2;
     0
     );
     P1 > 0
     );
     P1;
     """"
     )
     )
     ",
     * )
     */
    private function calculateDagenFoer()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_DAGEN_FOER, function () {
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $date_prev = $this->get(self::COLUMN_INPUT_DATE, -1);
            $date_prev_prev = $this->get(self::COLUMN_INPUT_DATE, -2);
            $_13_timer_prev = $this->get(self::COLUMN_TEMP_13_TIMER, -1);
            $_13_timer_prev_prev = $this->get(self::COLUMN_TEMP_13_TIMER, -2);

            if ($date - 1 === $date_prev && $_13_timer_prev > 0) {
                return $_13_timer_prev;
            } else {
                if ($date - 1 === $date_prev_prev && $_13_timer_prev_prev > 0) {
                    return $_13_timer_prev_prev;
                } else {
                    return '';
                }
            }
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     HVIS(
     OG(
     P3 = 0;
     Q3 = 0
     );
     0;
     HVIS(
     Q3 = 0;
     P3;
     HVIS(
     R3 = 0;
     HVIS(
     P3 = 0;
     Q3;
     P3 + Q3
     );
     HVIS(
     R3 >= Q3;
     P3;
     Q3 - R3
     )
     )
     )
     )
     ",
     * )
     */
    private function calculate100Pct()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_100_PCT, function () {
            $_13Timer = $this->calculate13Timer();
            $_11Timer = $this->calculate11Timer();
            $dagenFoer = $this->calculateDagenFoer();

            if (0 === $_13Timer and 0 === $_11Timer) {
                return 0;
            } else {
                if (0 === $_11Timer) {
                    return $_13Timer;
                } else {
                    if (0 === $dagenFoer) {
                        if (0 === $_13Timer) {
                            return $_11Timer;
                        } else {
                            return $_13Timer + $_11Timer;
                        }
                    } else {
                        if ($dagenFoer >= $_11Timer) {
                            return $_13Timer;
                        } else {
                            return $_11Timer - $dagenFoer;
                        }
                    }
                }
            }
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="=
     TÆL.HVIS(
     P21:Q21;
     "">0""
     )
     ",
     * )
     */
    private function calculateAntalHviletidsbrud()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_ANTAL_HVILETIDSBRUD, function () {
            $_13_timer = $this->calculate13Timer();
            $_11_timer = $this->calculate11Timer();

            return \count(array_filter([$_13_timer, $_11_timer], function ($value) {
                return $value > 0;
            }));
        });
    }

    /**
     * @var array
     */
    private $rows;

    /**
     * @var int
     */
    private $rowsIndex;

    private function setRows(array &$rows)
    {
        $this->rows = &$rows;
        $this->rowsIndex = (null === $rows || empty($rows)) ? null : -1;
    }

    private function nextRow()
    {
        if (!isset($this->rows)) {
            throw new \RuntimeException('No current rows');
        }
        if ($this->rowsIndex < \count($this->rows) - 1) {
            ++$this->rowsIndex;

            return true;
        }

        unset($this->rows);
        $this->rowsIndex = null;

        return false;
    }

    /**
     * Get a keyed value from row. Throw exception if key is not set.
     */
    private function get(string $key, int $offset = 0)
    {
        if (!isset($this->rows)) {
            throw new \RuntimeException('No current rows');
        }
        if (!\array_key_exists($this->rowsIndex, $this->rows)) {
            throw new \RuntimeException('No current row');
        }
        if (0 === $offset) {
            $row = $this->rows[$this->rowsIndex];
            // Require value in current row.
            if (!\array_key_exists($key, $row)) {
                throw new \RuntimeException(sprintf('Invalid row key: %s', $key));
            }

            return $row[$key];
        } else {
            return $this->rows[$this->rowsIndex + $offset][$key] ?? null;
        }
    }

    /**
     * Check if column is set in current row.
     */
    private function isSet(string $key)
    {
        if (null === $this->rows) {
            throw new \RuntimeException('No current rows');
        }
        if (!\array_key_exists($this->rowsIndex, $this->rows)) {
            throw new \RuntimeException('No current row');
        }

        return \array_key_exists($key, $this->rows[$this->rowsIndex]);
    }

    /**
     * Set value in current row.
     */
    private function set(string $key, $value)
    {
        $this->rows[$this->rowsIndex][$key] = $value;
    }

    private function testCheckRows(array $rows)
    {
        foreach ($rows as $index => $row) {
            foreach ([
                self::COLUMN_TEMP_TIMER => self::COLUMN_TEST_REFERENCE_TIMER,
            ] as $calculated => $reference) {
                if (!\array_key_exists($calculated, $row) || !\array_key_exists($reference, $row) || $row[$calculated] !== $row[$reference]) {
                    header('content-type: text/plain');
                    echo var_export([
                        $calculated => \array_key_exists($calculated, $row) ? $row[$calculated] : '👻',
                        $reference => \array_key_exists($reference, $row) ? $row[$reference] : '👻',
                        'row['.$index.']' => $row,
                    ], true);
                    die(__FILE__.':'.__LINE__.':'.__METHOD__);
                }
            }
        }
    }

    /**
     * Calculate and set column in a row.
     */
    private function calculateColumn(string $column, callable $calculate, $type = 'float')
    {
        if (!$this->isSet($column)) {
            $value = $calculate();
            if ('float' === $type) {
                $value = (float)$value;
            } elseif ('int' === $type) {
                $value = (int)$value;
            }

            $this->set($column, $value);
        }

        return $this->get($column);
    }
}
