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
     * @Setting(type="int", name="Timeløn"),
     *
     * @var int
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

    private const COLUMN_TEMP_OVERTID = 'overtid';
    private const COLUMN_TEMP_IS_OVERTIME = 'is overtime';
    private const COLUMN_TEMP_TIMER = 'timer';
    private const COLUMN_TEMP_NAT = 'nat';
    private const COLUMN_TEMP_IKKE_PLANLAGT7 = 'ikke planlagt 7';
    private const COLUMN_TEMP_13_TIMER = '13 timer';
    private const COLUMN_TEMP_11_TIMER = '11 timer';

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

                // Assume that actual end is on next day is less that actual start.
                if ($employeeRow[self::COLUMN_INPUT_ACTUAL_END] < $employeeRow[self::COLUMN_INPUT_ACTUAL_START]) {
                    $employeeRow[self::COLUMN_INPUT_ACTUAL_END] = $employeeRow[self::COLUMN_INPUT_ACTUAL_END]->add(new \DateInterval('P1D'));
                }

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

        $startDate = $this->dateTime2Excel($this->startDate);
        $endDate = $this->dateTime2Excel($this->endDate) + 1;

        foreach ($this->data as $employeeNumber => $rows) {
            // Do some calculations before trimming data.
            $this->calculateIsOvertime($rows);
            $this->calculateIkkePlanlagt7($rows);
            $this->calculate11Timer($rows);

            // Keep only rows in the specified report date interval.
            $rows = array_values(array_filter($rows, function (array $row) use ($startDate, $endDate) {
                return $startDate <= $row[self::COLUMN_INPUT_DATE]
                    && $row[self::COLUMN_INPUT_DATE] < $endDate;
            }));

            $this->result[$employeeNumber] = $this->calculateEmployee($rows);
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

        foreach ($this->result as $employeeNumber => $rows) {
            foreach ($rows as $row) {
                $row = [
                    self::COLUMN_INPUT_EMPLOYEE_NUMBER => $row[self::COLUMN_INPUT_EMPLOYEE_NUMBER],
                    self::COLUMN_INPUT_CONTRACT => $row[self::COLUMN_INPUT_CONTRACT],
                    self::COLUMN_INPUT_DATE => $this->formatExcelDate($row[self::COLUMN_INPUT_DATE]),
                    self::COLUMN_INPUT_EVENT => $row[self::COLUMN_INPUT_EVENT],
                    self::COLUMN_INPUT_PLANNED_START => $this->formatExcelTime($row[self::COLUMN_INPUT_PLANNED_START]),
                    self::COLUMN_INPUT_PLANNED_END => $this->formatExcelTime($row[self::COLUMN_INPUT_PLANNED_END]),
                    self::COLUMN_INPUT_ACTUAL_START => $this->formatExcelTime($row[self::COLUMN_INPUT_ACTUAL_START]),
                    self::COLUMN_INPUT_ACTUAL_END => $this->formatExcelTime($row[self::COLUMN_INPUT_ACTUAL_END]),

                    self::COLUMN_TEMP_TIMER => $this->formatExcelTime($row[self::COLUMN_TEMP_TIMER]),
                    self::COLUMN_TEMP_OVERTID => $this->formatExcelTime($row[self::COLUMN_TEMP_OVERTID]),
                    self::COLUMN_TEMP_NAT => $this->formatExcelTime($row[self::COLUMN_TEMP_NAT]),

                    self::COLUMN_TEMP_IS_OVERTIME => $row[self::COLUMN_TEMP_IS_OVERTIME],
                    self::COLUMN_TEMP_IKKE_PLANLAGT7 => $row[self::COLUMN_TEMP_IKKE_PLANLAGT7],
                ];

                if (1 === $rowIndex) {
                    $this->writeCells($sheet, 1, $rowIndex, array_keys($row));
                    ++$rowIndex;
                }
                $this->writeCells($sheet, 1, $rowIndex, $row);
                ++$rowIndex;
            }
            break;
        }

        return $result;

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

    private function calculateEmployee(array $rows)
    {
        foreach ($rows as $index => &$row) {
            $row[self::COLUMN_TEMP_OVERTID] = $this->calculateOvertid($row);
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
    // private const CONTRACT_TEKNIK_37_HOURS_3_MÅNEDER = 'Teknik 37 hours 3 måneder';
    private const CONTRACT_TEKNIK_37_HOURS = 'Teknik 37 hours';
    private const CONTRACT_TEKNIK_32_HOURS = 'Teknik 32 hours';
    private const CONTRACT_TIMELØNNEDE = 'Timelønnede';

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
    private function calculateTimer(array &$row)
    {
        $key = self::COLUMN_TEMP_TIMER;

        if (!\array_key_exists($key, $row)) {
            $calculate = function ($row) {
                $contract = $this->get(self::COLUMN_INPUT_CONTRACT, $row);
                $event = $this->get(self::COLUMN_INPUT_EVENT, $row);
                $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END, $row);

                if (self::EVENT_VAGT === $event
                    || (self::EVENT_SYGDOM === $event && !empty($actual_end))) {
                    return $row[self::COLUMN_INPUT_ACTUAL_END] - $row[self::COLUMN_INPUT_ACTUAL_START];
                } elseif (!$this->isNormnedsættende($event)) {
                    return 0;
                } else {
                    return $this->getUgenorm($contract) / 5 / 24;
                }
            };

            $row[$key] = $calculate($row);
        }

        return $row[$key];
    }

    private $normnedsaettendeItems;

    private function isNormnedsættende($event)
    {
        if (null === $this->normnedsaettendeItems) {
            $this->normnedsaettendeItems = array_filter(array_map('trim', explode(PHP_EOL, $this->normnedsaettende)));
        }

        return \in_array($event, $this->normnedsaettendeItems);
    }

    private $kontraktnormerItems;

    private function getUgenorm($contract)
    {
        if (null === $this->kontraktnormerItems) {
            $this->kontraktnormerItems = array_column(array_map('str_getcsv', array_filter(array_map('trim', explode(PHP_EOL, $this->kontraktnormer)))), null, 0);
        }

        if (!isset($this->kontraktnormerItems[$contract])) {
            throw new \RuntimeException(sprintf('Invalid contract: %s', $contract));
        }

        return $this->kontraktnormerItems[$contract][1];
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
    private function calculateIsOvertime(array &$rows)
    {
        $numberOfRows = \count($rows);
        foreach ($rows as $index => &$row) {
            $row[self::COLUMN_TEMP_IS_OVERTIME] =
                ($index > 0 && $row[self::COLUMN_INPUT_DATE] === $rows[$index - 1][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index - 1][self::COLUMN_INPUT_EVENT])
                || ($index > 1 && $row[self::COLUMN_INPUT_DATE] === $rows[$index - 2][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index - 2][self::COLUMN_INPUT_EVENT])
                || ($index < $numberOfRows - 1 && $row[self::COLUMN_INPUT_DATE] === $rows[$index + 1][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index + 1][self::COLUMN_INPUT_EVENT])
                || ($index < $numberOfRows - 2 && $row[self::COLUMN_INPUT_DATE] === $rows[$index + 2][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index + 2][self::COLUMN_INPUT_EVENT]);
        }
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
    private function calculateOvertid(array &$row)
    {
        $contract = $row[self::COLUMN_INPUT_CONTRACT];
        $planned_start = $row[self::COLUMN_INPUT_PLANNED_START];
        $planned_end = $row[self::COLUMN_INPUT_PLANNED_END];
        $actual_end = $row[self::COLUMN_INPUT_ACTUAL_END];
        $OT = $this->get(self::COLUMN_TEMP_IS_OVERTIME, $row);
        $timeløn = $this->timeloen;
        $timer = $this->calculateTimer($row);
        $date = $row[self::COLUMN_INPUT_DATE];
        $overtid_start = $this->overtidNatFra;
        $overtid_slut = $this->overtidNatTil;
        $nat = (int) $this->calculateNat($row);

        if (self::CONTRACT_TIMELØNNEDE !== $contract) {
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
                if ($actual_end > $planned_end && $timer > $timeløn) {
                    if ($actual_end < $date + 1) {
                        if ($planned_end - $planned_start < $timeløn) {
                            return $timer - $timeløn;
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
    private function calculateNat(array &$row)
    {
        $key = self::COLUMN_TEMP_NAT;

        if (!\array_key_exists($key, $row)) {
            $calculate = function ($row) {
                $event = $this->get(self::COLUMN_INPUT_EVENT, $row);
                $ikkePlanlagt7 = $this->get(self::COLUMN_TEMP_IKKE_PLANLAGT7, $row);
                $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START, $row);
                $actual_end = $this->get(self::COLUMN_INPUT_ACTUAL_END, $row);
                $date = $this->get(self::COLUMN_INPUT_DATE, $row);
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
            };

            $row[$key] = $calculate($row);
        }

        return $row[$key];
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
        E3 = E2;
        F2 = ""Løn: Ikke planlagt/7. dag""
    );
    J3 - I3;
    HVIS(
        OG(
            E3 = E1;
            F1 = ""Løn: Ikke planlagt/7. dag""
        );
        J3 - I3;
        HVIS(
            OG(
                E3 = E4;
                F4 = ""Løn: Ikke planlagt/7. dag""
            );
            J3 - I3;
            HVIS(
                OG(
                    E3 = E5;
                    F5 = ""Løn: Ikke planlagt/7. dag""
                );
                J3 - I3;
                0
            )
        )
    )
)
",
     * )
     */
    private function calculateIkkePlanlagt7(array &$rows)
    {
        $numberOfRows = \count($rows);

        foreach ($rows as $index => &$row) {
            $row[self::COLUMN_TEMP_IKKE_PLANLAGT7] =
                ($index > 0 && $row[self::COLUMN_INPUT_DATE] === $rows[$index - 1][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index - 1][self::COLUMN_INPUT_EVENT])
                || ($index > 1 && $row[self::COLUMN_INPUT_DATE] === $rows[$index - 2][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index - 2][self::COLUMN_INPUT_EVENT])
                || ($index < $numberOfRows - 1 && $row[self::COLUMN_INPUT_DATE] === $rows[$index + 1][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index + 1][self::COLUMN_INPUT_EVENT])
                || ($index < $numberOfRows - 2 && $row[self::COLUMN_INPUT_DATE] === $rows[$index + 2][self::COLUMN_INPUT_DATE] && self::EVENT_LØN_OVERTID === $rows[$index + 2][self::COLUMN_INPUT_EVENT]) ? 1 : 0;
        }
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
    private function calculate50Pct(array $row)
    {
        // If($event = ”Vagt” and sum("Overtid”;”Nat”;"Løn: Ikke planlagt/7. dag") > 0) {
// If($100% = 0) {
// return sum("Overtid”;”Nat”;"Løn: Ikke planlagt/7. dag")
// } else {
// If($100% >= sum("Overtid”;”Nat”;"Løn: Ikke planlagt/7. dag") {
// return 0
// } else {
// return sum("Overtid”;”Nat”;"Løn: Ikke planlagt/7. dag") - $100%
// }
// }
// } else {
// return 0
// }
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
    private function calculate13Timer(array $row)
    {
        if (self::EVENT_VAGT === $event and 0 !== $timer and $timer > $_13_timer) {
            return $timer - $_13_timer;
        } else {
            return 0;
        }
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
    private function calculate11Timer(array $rows)
    {
        $calculate = function ($row, $prevRow, $prevPrevRow) {
            $event = $this->get(self::COLUMN_INPUT_EVENT, $row);
            $date = $this->get(self::COLUMN_INPUT_DATE, $row);
            $date_prev = $prevRow[self::COLUMN_INPUT_DATE] ?? null;
            $date_prev_prev = $prevPrevRow[self::COLUMN_INPUT_DATE] ?? null;
            $event_prev = $prevRow[self::COLUMN_INPUT_EVENT] ?? null;
            $event_prev_prev = $prevPrevRow[self::COLUMN_INPUT_EVENT] ?? null;
            $actual_start = $this->get(self::COLUMN_INPUT_ACTUAL_START, $row);
            $actual_end_prev = $prevRow[self::COLUMN_INPUT_ACTUAL_END] ?? null;
            $actual_end_prev_prev = $prevPrevRow[self::COLUMN_INPUT_ACTUAL_END] ?? null;
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

        $key = self::COLUMN_TEMP_11_TIMER;
        foreach ($rows as $index => &$row) {
            $row[$key] = $calculate($row, $rows[$index - 1] ?? null, $rows[$index - 2] ?? null);
        }
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
    private function calculateDagenFør(array $row)
    {
        $key = self::COLUMN_TEMP_DAGEN_FØR;

        if (!\array_key_exists($key, $row)) {
            $calculate = function ($row) {
                $_13_timer_prev = $rows[$index - 1][self::COLUMN_TEMP_13_TIMER] ?? null;
                $_13_timer_prev_prev = $rows[$index - 2][self::COLUMN_TEMP_13_TIMER] ?? null;

                if ($date - 1 === $date_prev && $_13_timer_prev > 0) {
                    return $_13_timer_prev;
                } else {
                    if ($date - 1 === $date_prev_prev && $_13_timer_prev_prev > 0) {
                        return $_13_timer_prev_prev;
                    } else {
                        return '';
                    }
                }
            };

            $row[$key] = $calculate($row);
        }

        return $row[$key];
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
    private function calculate100Pct(array $row)
    {
        $key = self::COLUMN_TEMP_100_PCT;

        if (!\array_key_exists($key, $row)) {
            $calculate = function ($row) {
                if ($_13_timer and 0 === $_11_timer) {
                    return 0;
                } else {
                    if (0 === $_11_timer) {
                        return $_13_timer;
                    } else {
                        if (0 === $dagen_før) {
                            if (0 === $_13_timer) {
                                return $_11_timer;
                            } else {
                                return $_13_timer + $_11_timer;
                            }
                        } else {
                            if ($dagen_før >= $_11_timer) {
                                return $_13_timer;
                            } else {
                                return $_11_timer - $dagen_før;
                            }
                        }
                    }
                }
            };

            $row[$key] = $calculate($row);
        }

        return $row[$key];
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
    private function calculateAntalHviletidsbrud(array $row)
    {
        $key = self::COLUMN_TEMP_ANTAL_HVILETIDSBRUD;

        if (!\array_key_exists($key, $row)) {
            $calculate = function ($row) {
                $_13_timer = $this->get(self::COLUMN_TEMP_13_TIMER);
                $_11_timer = $this->get(self::COLUMN_TEMP_11_TIMER);

                return \count(array_filter([$_13_timer, $_11_timer], function ($value) {
                    return $value > 0;
                }));
            };

            $row[$key] = $calculate($row);
        }

        return $row[$key];
    }

    /**
     * Get a keyed value from row. Throw exception if key is not set.
     */
    private function get($key, array $row, bool $requireValue = true)
    {
        if ($requireValue && !\array_key_exists($key, $row)) {
            throw new \RuntimeException(sprintf('Invalid row key: %s', $key));
        }

        return $row[$key];
    }
}
