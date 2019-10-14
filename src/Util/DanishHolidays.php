<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Util;

use App\Calculator\Date;
use DateInterval;
use DateTimeInterface;
use DateTimeImmutable;

class DanishHolidays
{
    /**
     * Decide if the given day is a holiday.
     */
    public static function isHoliday(DateTimeInterface $date)
    {
        return null !== self::getHolidayName($date);
    }

    /**
     * Get name of holiday if any.
     */
    public static function getHolidayName(DateTimeInterface $date)
    {
        $year = (int) $date->format('Y');
        $holidays = self::getHolidays($year);
        $key = $date->format(DateTimeInterface::ATOM);

        return $holidays[$key] ?? null;
    }

    /**
     * Holidays indexed by year.
     *
     * @var array
     */
    private static $holidays;

    public static function getHolidays($year = null)
    {
        if (null === $year) {
            $year = (int) date('Y');
        }

        if (!isset(self::$holidays[$year])) {
            $march21 = new DateTimeImmutable($year.'-03-21');
            $easter = $march21->add(DateInterval::createFromDateString(easter_days($year).' days'));

            $holidays = [
                'nytårsdag' => new DateTimeImmutable($easter->format('Y-01-01')),
                'palmesøndag' => $easter->add(DateInterval::createFromDateString('-7 days')),
                'skærtorsdag' => $easter->add(DateInterval::createFromDateString('-3 days')),
                'langfredag' => $easter->add(DateInterval::createFromDateString('-2 days')),
                'påskedag' => $easter,
                '2. påskedag' => $easter->add(DateInterval::createFromDateString('1 day')),
                'store bededag' => $easter->add(DateInterval::createFromDateString('26 days')),
                'kristi himmelfartsdag' => $easter->add(DateInterval::createFromDateString('39 days')),
                'pinsedag' => $easter->add(DateInterval::createFromDateString('49 days')),
                '2. pinsedag' => $easter->add(DateInterval::createFromDateString('50 days')),
                'juledag' => new DateTimeImmutable($easter->format('Y-12-25')),
                '2. juledag' => new DateTimeImmutable($easter->format('Y-12-26')),
            ];

            foreach ($holidays as $name => $date) {
                self::$holidays[$year][$date->format(DateTimeInterface::ATOM)] = $date;
            }
        }

        return self::$holidays[$year];
    }
}
