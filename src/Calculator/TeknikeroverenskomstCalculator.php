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
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

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
     * @Setting(type="int", name="Standard", description="Antal arbejdstimer på en uge."),
     *
     * @var int
     */
    private $standard;

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

    private const EMPTY_VALUE = '';
    private const MISSING_VALUE = '(!!missing value!!)';

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

    private const COLUMN_OUTPUT_ARBEJDSUGE = 'Arbejdsuge';
    private const COLUMN_OUTPUT_NORM = 'Norm';
    private const COLUMN_OUTPUT_MERTID = 'Mertid';
    private const COLUMN_OUTPUT_AFSPADSERING = 'Afspadsering';
    private const COLUMN_OUTPUT_DELTIDSFRADRAG = 'Deltidsfradrag';

    private const COLUMN_SUM_P_5571 = 'Σ P 5571';
    private const COLUMN_SUM_P_6625 = 'Σ P 6625';
    private const COLUMN_SUM_P_MILJOE = 'Σ P Miljø';
    private const COLUMN_SUM_P_VARSEL = 'Σ P Varsel';
    private const COLUMN_SUM_P_DELT = 'Σ P Delt';
    private const COLUMN_SUM_P_50_PCT = 'Σ P 50%';
    private const COLUMN_SUM_P_100_PCT = 'Σ P 100%';
    private const COLUMN_SUM_P_ANTAL = 'Σ P Antal';
    private const COLUMN_SUM_P_NORMAL = 'Σ P Normal';
    private const COLUMN_SUM_TIMER2 = 'Σ Timer2';
    private const COLUMN_SUM_ARBEJDSDAGE = 'Σ Arbejdsdage';

    private const COLUMN_OUTPUT_EMPLOYEE_NUMBER = 'Medarbejdernummer';
    private const COLUMN_OUTPUT_LOENART = 'Lønart';
    private const COLUMN_OUTPUT_LOEBENR = 'Løbenr.';
    private const COLUMN_OUTPUT_ENHEDER_I_ALT = 'Enheder (i alt)';
    private const COLUMN_OUTPUT_IKRAFT_DATO = 'Ikraft dato (for lønmåned)';

    private const COLUMN_TEMP_OVERTID = 'overtid';
    private const COLUMN_TEMP_VARSEL = 'varsel';
    private const COLUMN_TEMP_OT = 'is overtime';
    private const COLUMN_TEMP_TIMER = 'timer';
    private const COLUMN_TEMP_NAT = 'nat';
    private const COLUMN_TEMP_IKKE_PLANLAGT_7 = 'ikke planlagt 7';
    private const COLUMN_TEMP_13_TIMER = '13 timer';
    private const COLUMN_TEMP_11_TIMER = '11 timer';
    private const COLUMN_TEMP_5571 = '5571';
    private const COLUMN_TEMP_6625 = '6625';
    private const COLUMN_TEMP_MILJOE = 'miljø';
    private const COLUMN_TEMP_50_PCT = '50 %';
    private const COLUMN_TEMP_100_PCT = '100 %';
    private const COLUMN_TEMP_DAGEN_FOER = 'dagen før';
    private const COLUMN_TEMP_ANTAL_HVILETIDSBRUD = 'antal hviletidsbrud';
    private const COLUMN_TEMP_DELT = 'Delt';

    private const COLUMN_TEST_REFERENCE_TIMER = 'test reference Timer';
    private const COLUMN_TEST_REFERENCE_OVERTID = 'test reference Overtid';
    private const COLUMN_TEST_REFERENCE_NAT = 'test reference Nat';
    private const COLUMN_TEST_REFERENCE_IKKE_PLANLAGT_7 = 'test reference Ikke pl./7';
    private const COLUMN_TEST_REFERENCE_50_PCT = 'test reference 50%';
    private const COLUMN_TEST_REFERENCE_13_TIMER = 'test reference 13 timer';
    private const COLUMN_TEST_REFERENCE_11_TIMER = 'test reference 11 timer';
    private const COLUMN_TEST_REFERENCE_DAGEN_FOER = 'test reference Dagen før';
    private const COLUMN_TEST_REFERENCE_100_PCT = 'test reference 100%';
    private const COLUMN_TEST_REFERENCE_ANTAL_HVILETIDSBRUD = 'test reference Antal Hv.';
    private const COLUMN_TEST_REFERENCE_5571 = 'test reference 5571';
    private const COLUMN_TEST_REFERENCE_6625 = 'test reference 6625';
    private const COLUMN_TEST_REFERENCE_MILJOE = 'test reference Miljø';
    private const COLUMN_TEST_REFERENCE_VARSEL = 'test reference Varsel';
    private const COLUMN_TEST_REFERENCE_DELT = 'test reference Delt';
    // private const COLUMN_TEST_REFERENCE_HELLIGDAG = 'test reference Helligdag';
    private const COLUMN_TEST_REFERENCE_OT = 'test reference OT';
    private const COLUMN_TEST_REFERENCE_P_5571 = 'test reference P 5571';
    private const COLUMN_TEST_REFERENCE_P_6625 = 'test reference P 6625';
    private const COLUMN_TEST_REFERENCE_P_MILJOE = 'test reference P Miljø';
    private const COLUMN_TEST_REFERENCE_P_VARSEL = 'test reference P Varsel';
    private const COLUMN_TEST_REFERENCE_P_DELT = 'test reference P Delt';
    private const COLUMN_TEST_REFERENCE_P_50_PCT = 'test reference P 50%';
    private const COLUMN_TEST_REFERENCE_P_100_PCT = 'test reference P 100%';
    private const COLUMN_TEST_REFERENCE_P_ANTAL = 'test reference P Antal';
    private const COLUMN_TEST_REFERENCE_P_NORMAL = 'test reference P Normal';
    private const COLUMN_TEST_REFERENCE_TIMER2 = 'test reference Timer 2';
    private const COLUMN_TEST_REFERENCE_ARBEJDSDAGE = 'test reference Arbejdsdage';

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
                        self::COLUMN_TEST_REFERENCE_DELT => $row[self::EXCEL_COLUMN_Y],
                        // self::COLUMN_TEST_REFERENCE_HELLIGDAG => $row[self::EXCEL_COLUMN_Z],
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
        if ($this->standard <= 0) {
            throw new InvalidArgumentException('Standard must be a positive integer');
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
        $this->startDate = $this->dateTime2Excel($this->startDate);
        $this->endDate = $this->dateTime2Excel($this->endDate) + 1;

        foreach ($this->data as $employeeNumber => &$rows) {
            $this->calculateEmployee($rows);

            if ($this->testMode) {
                $this->testCheckRows($rows);
            }

            // Pivot
            // Group by Employee name, number, contract, email
            $employee = [];
            foreach ($rows as $row) {
                $key = implode('|||', [$row[self::COLUMN_INPUT_NAME], $row[self::COLUMN_INPUT_EMPLOYEE_NUMBER], $row[self::COLUMN_INPUT_CONTRACT], $row[self::COLUMN_INPUT_EMAIL]]);
                if (!isset($employee[$key])) {
                    $employee[$key] = [
                        self::COLUMN_INPUT_NAME => $row[self::COLUMN_INPUT_NAME],
                        self::COLUMN_INPUT_EMPLOYEE_NUMBER => $row[self::COLUMN_INPUT_EMPLOYEE_NUMBER],
                        self::COLUMN_INPUT_CONTRACT => $row[self::COLUMN_INPUT_CONTRACT],
                        self::COLUMN_INPUT_EMAIL => $row[self::COLUMN_INPUT_EMAIL],
                        self::COLUMN_SUM_P_5571 => 0,
                        self::COLUMN_SUM_P_6625 => 0,
                        self::COLUMN_SUM_P_MILJOE => 0,
                        self::COLUMN_SUM_P_VARSEL => 0,
                        self::COLUMN_SUM_P_DELT => 0,
                        self::COLUMN_SUM_P_50_PCT => 0,
                        self::COLUMN_SUM_P_100_PCT => 0,
                        // self::COLUMN_SUM_P_ANTAL => 0,
                        self::COLUMN_SUM_P_NORMAL => 0,
                        // self::COLUMN_SUM_TIMER2 => 0,
                        // self::COLUMN_SUM_ARBEJDSDAGE => 0,
                    ];
                }
                foreach ([
                    self::COLUMN_SUM_P_5571 => self::COLUMN_OUTPUT_P_5571,
                    self::COLUMN_SUM_P_6625 => self::COLUMN_OUTPUT_P_6625,
                    self::COLUMN_SUM_P_MILJOE => self::COLUMN_OUTPUT_P_MILJOE,
                    self::COLUMN_SUM_P_VARSEL => self::COLUMN_OUTPUT_P_VARSEL,
                    self::COLUMN_SUM_P_DELT => self::COLUMN_OUTPUT_P_DELT,
                    self::COLUMN_SUM_P_50_PCT => self::COLUMN_OUTPUT_P_50_PCT,
                    self::COLUMN_SUM_P_100_PCT => self::COLUMN_OUTPUT_P_100_PCT,
                    // self::COLUMN_SUM_P_ANTAL => self::COLUMN_OUTPUT_P_ANTAL,
                    self::COLUMN_SUM_P_NORMAL => self::COLUMN_OUTPUT_P_NORMAL,
                    // self::COLUMN_SUM_TIMER2 => self::COLUMN_OUTPUT_TIMER2,
                    // self::COLUMN_SUM_ARBEJDSDAGE => self::COLUMN_OUTPUT_ARBEJDSDAGE,
                ] as $sum => $column) {
                    $employee[$key][$sum] += self::EMPTY_VALUE === $row[$column] ? 0.0 : $row[$column];
                }
            }

            foreach ($employee as &$row) {
                $contract = $row[self::COLUMN_INPUT_CONTRACT];
                $row[self::COLUMN_OUTPUT_ARBEJDSUGE] = $this->getUgenorm($contract);
                $row[self::COLUMN_OUTPUT_NORM] = $row[self::COLUMN_OUTPUT_ARBEJDSUGE] / 5 * $this->getNumberOfWorkdays($this->getNormperiode($contract));
                $row[self::COLUMN_OUTPUT_MERTID] = $row[self::COLUMN_SUM_P_NORMAL] - $row[self::COLUMN_OUTPUT_NORM];
                $row[self::COLUMN_OUTPUT_DELTIDSFRADRAG] =
                    $row[self::COLUMN_OUTPUT_ARBEJDSUGE] / $this->standard
                    ? $row[self::COLUMN_OUTPUT_NORM] / $row[self::COLUMN_OUTPUT_ARBEJDSUGE] * ($this->standard - $row[self::COLUMN_OUTPUT_ARBEJDSUGE])
                    : 0;
                $row[self::COLUMN_OUTPUT_AFSPADSERING] = $row[self::COLUMN_SUM_P_50_PCT] * 1.5
                    + $row[self::COLUMN_SUM_P_100_PCT] * 2
                    + (
                        $row[self::COLUMN_OUTPUT_MERTID] - $row[self::COLUMN_OUTPUT_DELTIDSFRADRAG] > 0
                        ? ($row[self::COLUMN_OUTPUT_MERTID] - $row[self::COLUMN_OUTPUT_DELTIDSFRADRAG]) * 1.5
                        : $row[self::COLUMN_OUTPUT_MERTID]
                    );
            }

            foreach ($employee as $row) {
                $this->result[] = $row;
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

        $columns = [
            self::COLUMN_SUM_P_NORMAL,
            self::COLUMN_SUM_P_50_PCT,
            self::COLUMN_SUM_P_100_PCT,
            self::COLUMN_SUM_P_5571,
            self::COLUMN_SUM_P_6625,
            self::COLUMN_SUM_P_DELT,
            self::COLUMN_SUM_P_VARSEL,
            self::COLUMN_SUM_P_MILJOE,
            self::COLUMN_OUTPUT_AFSPADSERING,

            // self::COLUMN_SUM_P_ANTAL,
            // self::COLUMN_SUM_TIMER2,
            // self::COLUMN_SUM_ARBEJDSDAGE,
        ];

        $iKraftDato = $this->dateTime2Excel(new DateTimeImmutable($this->getExcelDate($this->endDate - 1)->format(DateTimeInterface::ATOM).' last day of month'));

        usort($this->result, function ($a, $b) {
            return strcmp($a[self::COLUMN_INPUT_EMPLOYEE_NUMBER], $b[self::COLUMN_INPUT_EMPLOYEE_NUMBER]);
        });

        foreach ($this->result as $row) {
            $contract = $row[self::COLUMN_INPUT_CONTRACT];
            foreach ($columns as $column) {
                if (('Timelønnede' === $contract && \in_array($column, [self::COLUMN_OUTPUT_AFSPADSERING], true))
                    || ('Timelønnede' !== $contract && \in_array($column, [self::COLUMN_SUM_P_NORMAL, self::COLUMN_SUM_P_MILJOE, self::COLUMN_SUM_P_50_PCT, self::COLUMN_SUM_P_100_PCT], true))) {
                    continue;
                }
                if (isset($row[$column]) && $row[$column] > 0) {
                    [$loenart, $loebenr] = $this->getLoenartAndLoebenr($column, $contract);
                    $this->writeCells($sheet, 1, $rowIndex, [
                        $row[self::COLUMN_INPUT_EMPLOYEE_NUMBER],
                        $loenart,
                        $loebenr,
                        $row[$column],
                        $iKraftDato,
                    ]);
                    $sheet->getStyleByColumnAndRow(4, $rowIndex)
                        ->getNumberFormat()
                        ->setFormatCode(NumberFormat::FORMAT_NUMBER_00);
                    $sheet->getStyleByColumnAndRow(5, $rowIndex)
                        ->getNumberFormat()
                        ->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
                    ++$rowIndex;
                }
            }
        }

        return $result;
    }

    private function getLoenartAndLoebenr(string $column, string $contract)
    {
        switch ($column) {
            case self::COLUMN_SUM_P_MILJOE:
                return ['0424', '1'];
            case self::COLUMN_SUM_P_NORMAL:
                return ['0140', '1'];
            case self::COLUMN_SUM_P_5571:
                return ['0557', '1'];
            case self::COLUMN_SUM_P_6625:
                return ['0662', '5'];
            case self::COLUMN_SUM_P_VARSEL:
                return ['0747', '1'];
            case self::COLUMN_SUM_P_DELT:
                return ['0379', '1'];
            case self::COLUMN_SUM_P_50_PCT:
                return ['0104', '1'];
            case self::COLUMN_SUM_P_100_PCT:
                return ['0109', '1'];
            case self::COLUMN_OUTPUT_AFSPADSERING:
                return ['0120', '1'];
        }

        throw new \RuntimeException(sprintf('Unknown column: %s', $column));
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
            $this->calculateTimer2();
            $this->calculateArbejdsdage();
        }

        return $rows;
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
     * =IF (ISERROR(VLOOKUP(E3, Helligdage!$B:$B, 1, FALSE)),
     *     IF (F3 = ""Vagt"",
     *           IF (WEEKDAY(E3, 2) < 6,
     *                  IF (I3 < E3 + Meta!$D$4,
     *                       IF (J3 < E3 + Meta!$D$4,
     *                              J3 - I3,
     *                              IF (J3 < E3 + Meta!$D$3,
     *                                   E3 + Meta!$D$4 - I3,
     *                                   E3 + Meta!$D$4 - I3 + J3 - E3 + Meta!$D$3
     *                              )
     *                       ),
     *                       IF (J3 < E3 + Meta!$D$3,
     *                              0,
     *                              IF (J3 < E3 + 1 + Meta!$D$4,
     *                                   J3 - E3 - Meta!$D$3,
     *                                   1 + Meta!$D$4 - I3
     *                              )
     *                       )
     *                  ),
     *                  IF (WEEKDAY(E3, 2) = 6,
     *                       IF (I3 < E3 + Meta!$D$4,
     *                              IF (J3 < E3 + Meta!$D$4,
     *                                   J3 - I3,
     *                                   0
     *                              ),
     *                              IF (J3 < E3 + Meta!$D$2,
     *                                   J3 - I3,
     *                                   IF (I3 < E3 + Meta!$D$2,
     *                                          IF (J3 < E3 + Meta!$D$3,
     *                                               J3 - ( E3 + Meta!$D$2 ),
     *                                               E3 + Meta!$D$3 - ( E3 + Meta!$D$2 )
     *                                          ),
     *                                          IF (J3 < E3 + Meta!$D$3,
     *                                               J3 - I3,
     *                                               E3 + Meta!$D$3 - I3
     *                                          )
     *                                   )
     *                              )
     *                       ),
     *                       0
     *                  )
     *           ),
     *           0
     *     ),
     *     0
     * )
     * "
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
                            return $actual_end - $actual_start;
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
            $_5571 = $this->calculate5571();
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            return 24 * $_5571;
        }, 'mixed');
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
     * =IF( Z3 = ""X"",
     *   IF(OR(
     *       J3 < E3 + 1,
     *       WEEKDAY(E3 + 1, 2) = 7,
     *       OFFSET(Z3, 1, 0) = ""X""
     *     ),
     *     J3 - I3,
     *     IF(
     *       WEEKDAY(E3 + 1, 2) = 1,
     *       IF(
     *         J3 < E3 + 1 + Meta!$E$3,
     *         J3 - I3,
     *         E3 + 1 + Meta!$E$3
     *       ),
     *       0
     *     )
     *   ),
     *   IF(
     *     F3 = ""Vagt"",
     *     IF(
     *       OR(
     *         WEEKDAY(E3, 2) >= 6,
     *         WEEKDAY(E3, 2) = 1
     *       ),
     *       IF(
     *         WEEKDAY(E3, 2) = 6,
     *         IF(
     *           I3 < E3 + Meta!$E$2,
     *           IF(
     *             J3 <= E3 + Meta!$E$2,
     *             0,
     *             J3 - ( E3 + Meta!$E$2 )
     *           ),
     *           J3 - I3
     *         ),
     *         IF(
     *           WEEKDAY(E3, 2) = 7,
     *           IF(
     *             J3 < E3 + 1 + Meta!$E$3,
     *             J3 - I3,
     *             E3 + 1 + Meta!$E$3 - I3
     *           ),
     *           IF(
     *             I3 > E3 + Meta!$E$3,
     *             0,
     *             IF(
     *               J3 > E3 + Meta!$E$3,
     *               E3 + Meta!$E$3 - I3,
     *               J3 - I3
     *             )
     *           )
     *         )
     *       ),
     *       0
     *     ),
     *     0
     *   )
     * )
     * "
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
                                    return 0;
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
                                if ($actual_start > $date + $_6625_slut) {
                                    return 0;
                                } else {
                                    if ($actual_end > $date + $_6625_slut) {
                                        return $date + $_6625_slut - $actual_start;
                                    } else {
                                        return $actual_end - $actual_start;
                                    }
                                }
                            }

                            return 0;
                        }
                    }
                }

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
     *     excelFormula=""
     * )
     */
    private function calculateP6625()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_6625, function () {
            $_6625 = $this->calculate6625();
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            return 24 * $_6625;
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
            $mijoe = $this->calculateMiljoe();
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            return 24 * $mijoe;
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="=HVIS($F3=""Løn: Varsel"";1;"""")"
     * )
     */
    private function calculateVarsel()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_VARSEL, function () {
            return self::EVENT_LOEN_VARSEL === $this->get(self::COLUMN_INPUT_EVENT) ? 1 : self::EMPTY_VALUE;
        }, 'mixed');
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
    private function calculatePVarsel()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_VARSEL, function () {
            return $this->calculateVarsel();
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="=HVIS($F3=""Løn: Delt tjeneste"";1;"""")"
     * )
     *
     * @return bool
     */
    private function calculateDelt()
    {
        return $this->calculateColumn(self::COLUMN_TEMP_DELT, function () {
            return self::EVENT_LOEN_DELT_TJENESTE === $this->get(self::COLUMN_INPUT_EVENT) ? 1 : self::EMPTY_VALUE;
        }, 'mixed');
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
            return $this->calculateDelt();
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
     HVIS(OG(F3 = ""Vagt""; SUM(L3:N3) > 0);
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
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="
     * =IF(
     *     AND(
     *         HLOOKUP(
     *             VLOOKUP(
     *                 Data!C3,
     *                 Overenskomst,
     *                 3,
     *                 FALSE
     *             ),
     *             Periode,
     *             2,
     *             FALSE
     *         ) <= Data!$E3,
     *         Data!$E3 <=
     *         HLOOKUP(
     *             VLOOKUP(
     *                 Data!C3,
     *                 Overenskomst,
     *                 3,
     *                 FALSE
     *             ),
     *             Periode,
     *             3,
     *             FALSE
     *         )
     *     ),
     *     O3 * 24,
     *     """"
     * )
     * "
     * )
     */
    private function calculateP50Pct()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_P_50_PCT, function () {
            $_50Pct = $this->calculate50Pct();
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            return 24 * $_50Pct;
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
            $_100Pct = $this->calculate100Pct();
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            return 24 * $_100Pct;
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
            $antal = $this->calculateAntalHviletidsbrud();
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            return $antal;
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
                    return 0;
                }
            } else {
                return self::EMPTY_VALUE;
            }
        });
    }

    /**
     * @Calculation(
     *     name="arbejdstimer",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula="=HVIS(OG(VOPSLAG(LOPSLAG(Data!C3;Overenskomst;3;FALSK);Periode;2;FALSK)<=Data!$E3;Data!$E3<=Slutdato);K3*24;"""")"
     * )
     */
    private function calculateTimer2()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_TIMER2, function () {
            $date = $this->get(self::COLUMN_INPUT_DATE);
            $contract = $this->get(self::COLUMN_INPUT_CONTRACT);
            $normperiode = $this->getNormperiode($contract);

            if ($normperiode[0] <= $date && $date <= $normperiode[1]) {
                return $this->calculateTimer() * 24;
            } else {
                return self::EMPTY_VALUE;
            }
        });
    }

    /**
     * Er der blevet arbejdet?
     *
     * @Calculation(
     *     name="arbejdsdage",
     *     description="",
     *     formula="",
     *     overenskomsttekst="",
     *     excelFormula=""
     * )
     *
     * @return int
     */
    private function calculateArbejdsdage()
    {
        return $this->calculateColumn(self::COLUMN_OUTPUT_ARBEJDSDAGE, function () {
            if (!$this->includeRow()) {
                return self::EMPTY_VALUE;
            }

            $date = $this->get(self::COLUMN_INPUT_DATE);
            $datePrev = $this->get(self::COLUMN_INPUT_DATE, -1);

            return (null !== $datePrev && $date !== $datePrev) ? 1 : self::EMPTY_VALUE;
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
            $startDate = $this->getExcelDate($this->startDate);

            switch ($period) {
                case 0:
                    return [
                        new DateTimeImmutable('2001-01-01'),
                        new DateTimeImmutable('2000-01-01'),
                    ];
                case 1:
                    $offset = $startDate->format(DateTimeInterface::ATOM);

                    return [
                        new DateTimeImmutable($offset.' first day of month'),
                        new DateTimeImmutable($offset.' last day of month'),
                    ];
                case 3:
                    // Get quarter containing start date.
                    $month = (int) $startDate->format('n');
                    $startQuarterMonth = 3 * (int) floor(($month - 1) / 3) + 1;

                    return [
                        new DateTimeImmutable($startDate->format(sprintf('Y-%02d-d\TH:i:sP', $startQuarterMonth)).' first day of month'),
                        new DateTimeImmutable($startDate->format(sprintf('Y-%02d-d\TH:i:sP', $startQuarterMonth + 2)).' last day of month'),
                    ];

                default:
                    throw new \RuntimeException(sprintf('Invalid norm period: %d', $period));
            }
        };

        $norm = $this->kontraktnormer($contract);
        $result = $calculate((int) $norm[2]);

        if ($asExcelDates) {
            $result = array_map([$this, 'dateTime2Excel'], $result);
        }

        return $result;
    }

    /**
     * Get number of workdays in a period.
     */
    private function getNumberOfWorkdays(array $period)
    {
        $numberOfWorkdays = 0;

        for ($d = $period[0]; $d <= $period[1]; ++$d) {
            $weekday = $this->getWeekday($d);
            if (self::WEEKDAY_SATURDAY === $weekday || self::WEEKDAY_SUNDAY === $weekday || $this->isHoliday($d)) {
                continue;
            }

            ++$numberOfWorkdays;
        }

        return $numberOfWorkdays;
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
     *     excelFormula="
=HVIS(
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
        return $this->calculateColumn(self::COLUMN_TEMP_OT, function () {
            return ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -2) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, -2))
                || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -1) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, -1))
                || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +1) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, +1))
                || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +2) && self::EVENT_LOEN_OVERTID === $this->get(self::COLUMN_INPUT_EVENT, +2)) ? 'OT' : self::EMPTY_VALUE;
        }, 'mixed');
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
     *     excelFormula="
     * =IF(
     *     C3 <> ""Timelønnede"",
     *     IF(
     *         AND(
     *             J3 > H3,
     *             N3 = 0,
     *             AA3 = ""OT""
     *         ),
     *         IF(
     *             J3 < E3 + 1,
     *             J3 - H3,
     *             IF(
     *                 H3 >= E3 + 1,
     *                 0,
     *                 E3 + Meta!$A$2 - H3
     *             )
     *         ),
     *         0
     *     ),
     *     IF(
     *         AND(
     *             J3 > H3,
     *             K3 > Meta!$G$2
     *         ),
     *         IF(
     *             J3 < E3 + 1,
     *             IF(
     *                 H3 - G3 < Meta!$G$2,
     *                 K3 - Meta!$G$2,
     *                 K3 - ( H3 - G3 )
     *             ),
     *             IF(
     *                 J3 < E3 + 1 + Meta!$A$3,
     *                 IF(
     *                     H3 - G3 < Meta!$G$2,
     *                     J3 - Meta!$G$2 - M3,
     *                     K3 - ( H3 - G3 ) - M3
     *                 ),
     *                 IF(
     *                     I3 < E3 + 21 / 24,
     *                     0,
     *                     J3 - ( E3 + 1 + Meta!$A$3 )
     *                 )
     *             )
     *         ),
     *         0
     *     )
     * )
     * "
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
            // HACK!
            if (0 === $overtid_start) {
                ++$overtid_start;
            }
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
                            return $date + $overtid_start - $planned_end;
                        }
                    }
                } else {
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
                } else {
                    return 0;
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
     *     excelFormula="
     =HVIS(
     OG(F7 = ""Vagt""; N7 = 0);
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
            $overtid_slut = $this->overtidNatTil;

            if (self::EVENT_VAGT === $event && 0 === $ikkePlanlagt7) {
                if ($actual_start < $date + $overtid_slut) {
                    if ($actual_end < $date + $overtid_slut) {
                        return $actual_end - $actual_start;
                    } else {
                        return $date + $overtid_slut - $actual_start;
                    }
                } else {
                    if ($actual_end < $date + 1) {
                        return 0;
                    } else {
                        if ($actual_end < $date + 1 + $overtid_slut) {
                            return $actual_end - ($date + 1);
                        } else {
                            return $overtid_slut;
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
        return $this->calculateColumn(self::COLUMN_TEMP_IKKE_PLANLAGT_7, function () {
            return (($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -2) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, -2))
                    || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, -1) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, -1))
                    || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +1) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, +1))
                    || ($this->get(self::COLUMN_INPUT_DATE) === $this->get(self::COLUMN_INPUT_DATE, +2) && self::EVENT_LOEN_IKKE_PLANLAGT_7_DAG === $this->get(self::COLUMN_INPUT_EVENT, +2)))
                ? $this->get(self::COLUMN_INPUT_ACTUAL_END) - $this->get(self::COLUMN_INPUT_ACTUAL_START) : 0;
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

            if (self::EVENT_VAGT === $event and 0 !== $timer and $timer > $_13_timer + 0.01) {
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
        return $this->calculateColumn(self::COLUMN_TEMP_11_TIMER, function () {
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
        });
    }

    /**
     * @Calculation(
     *     name="",
     *     description="",
     *     formula="",
     *     placeholders={},
     *     overenskomsttekst="",
     *     excelFormula="
     =HVIS(
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
                    return 0;
                }
            }
        }, 'mixed');
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
                        // @TODO Apparently the empty (or any?) string is greater than or equal to a number!
                        if (self::EMPTY_VALUE === $dagenFoer || $dagenFoer >= $_11Timer) {
                            return $_13Timer;
                        } else {
                            return $_11Timer - (float) $dagenFoer;
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

    private function includeRow()
    {
        $date = $this->get(self::COLUMN_INPUT_DATE);

        return $this->startDate <= $date && $date < $this->endDate;
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

    private function currentRow()
    {
        if (!\array_key_exists($this->rowsIndex, $this->rows)) {
            throw new \RuntimeException('No current row');
        }

        return $this->rows[$this->rowsIndex];
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

    /**
     * Calculate and set column in a row.
     */
    private function calculateColumn(string $column, callable $calculate, $type = 'mixed')
    {
        if (!$this->isSet($column)) {
            $value = $calculate();
            if (null !== $value) {
                if ('float' === $type) {
                    $value = (float) $value;
                } elseif ('int' === $type) {
                    $value = (int) $value;
                }
            }

            $this->set($column, $value);
        }

        return $this->get($column);
    }

    private function testCheckRows(array $rows)
    {
        return;
        foreach ($rows as $index => $row) {
            foreach ([
                self::COLUMN_TEMP_TIMER => self::COLUMN_TEST_REFERENCE_TIMER,
                self::COLUMN_TEMP_OVERTID => self::COLUMN_TEST_REFERENCE_OVERTID,
                self::COLUMN_TEMP_NAT => self::COLUMN_TEST_REFERENCE_NAT,
                self::COLUMN_TEMP_IKKE_PLANLAGT_7 => self::COLUMN_TEST_REFERENCE_IKKE_PLANLAGT_7,
                self::COLUMN_TEMP_100_PCT => self::COLUMN_TEST_REFERENCE_100_PCT,
                self::COLUMN_TEMP_50_PCT => self::COLUMN_TEST_REFERENCE_50_PCT,
                self::COLUMN_TEMP_13_TIMER => self::COLUMN_TEST_REFERENCE_13_TIMER,
                self::COLUMN_TEMP_11_TIMER => self::COLUMN_TEST_REFERENCE_11_TIMER,
                self::COLUMN_TEMP_DAGEN_FOER => self::COLUMN_TEST_REFERENCE_DAGEN_FOER,
                self::COLUMN_TEMP_ANTAL_HVILETIDSBRUD => self::COLUMN_TEST_REFERENCE_ANTAL_HVILETIDSBRUD,
                self::COLUMN_TEMP_5571 => self::COLUMN_TEST_REFERENCE_5571,
                self::COLUMN_TEMP_6625 => self::COLUMN_TEST_REFERENCE_6625,
                self::COLUMN_TEMP_MILJOE => self::COLUMN_TEST_REFERENCE_MILJOE,
                // self::COLUMN_TEMP_VARSEL => self::COLUMN_TEST_REFERENCE_VARSEL,
                self::COLUMN_TEMP_DELT => self::COLUMN_TEST_REFERENCE_DELT,
                // self::COLUMN_TEMP_HELLIGDAG => self::COLUMN_TEST_REFERENCE_HELLIGDAG,
                self::COLUMN_TEMP_OT => self::COLUMN_TEST_REFERENCE_OT,
                self::COLUMN_OUTPUT_P_5571 => self::COLUMN_TEST_REFERENCE_P_5571,
                self::COLUMN_OUTPUT_P_6625 => self::COLUMN_TEST_REFERENCE_P_6625,
                self::COLUMN_OUTPUT_P_MILJOE => self::COLUMN_TEST_REFERENCE_P_MILJOE,
                self::COLUMN_OUTPUT_P_VARSEL => self::COLUMN_TEST_REFERENCE_P_VARSEL,
                self::COLUMN_OUTPUT_P_DELT => self::COLUMN_TEST_REFERENCE_P_DELT,
                self::COLUMN_OUTPUT_P_50_PCT => self::COLUMN_TEST_REFERENCE_P_50_PCT,
                self::COLUMN_OUTPUT_P_100_PCT => self::COLUMN_TEST_REFERENCE_P_100_PCT,
                self::COLUMN_OUTPUT_P_ANTAL => self::COLUMN_TEST_REFERENCE_P_ANTAL,
                self::COLUMN_OUTPUT_P_NORMAL => self::COLUMN_TEST_REFERENCE_P_NORMAL,
                self::COLUMN_OUTPUT_TIMER2 => self::COLUMN_TEST_REFERENCE_TIMER2,
                self::COLUMN_OUTPUT_ARBEJDSDAGE => self::COLUMN_TEST_REFERENCE_ARBEJDSDAGE,
            ] as $calculated => $reference) {
                if (!\array_key_exists($calculated, $row) || !\array_key_exists($reference, $row) || !$this->testEquals($row[$calculated], $row[$reference])) {
                    header('content-type: text/plain');
                    echo var_export([
                        $calculated => \array_key_exists($calculated, $row) ? $row[$calculated] : self::MISSING_VALUE,
                        $reference => \array_key_exists($reference, $row) ? $row[$reference] : self::MISSING_VALUE,
                        'row['.$index.']' => array_merge($row, [
                            self::COLUMN_INPUT_DATE => $this->formatExcelDate($row[self::COLUMN_INPUT_DATE]),
                            self::COLUMN_INPUT_ACTUAL_START => $this->formatExcelTime($row[self::COLUMN_INPUT_ACTUAL_START]),
                            self::COLUMN_INPUT_ACTUAL_END => $this->formatExcelTime($row[self::COLUMN_INPUT_ACTUAL_END]),
                        ]),
                    ], true);
                    die(__FILE__.':'.__LINE__.':'.__METHOD__);
                }
            }
        }
    }

    private $delta = 0.1e-8;

    private function testEquals($operand1, $operand2)
    {
        if (is_numeric($operand1) && is_numeric($operand2)) {
            return abs($operand1 - $operand2) < $this->delta;
        } else {
            return $operand1 === $operand2;

            return 0 === strcmp($operand1, $operand2);
        }
    }
}
