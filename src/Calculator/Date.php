<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Calculator;

use DateInterval;
use DateTimeImmutable;
use DateTimeInterface;

/**
 * Helper class to make date calculation (a little) easier.
 *
 * Wraps DateTimeImmutable.
 */
class Date extends DateTimeImmutable
{
    /**
     * Create a new Date from a DateTime.
     *
     * @param DateTimeInterface $date
     *
     * @return static
     *
     * @throws \Exception
     */
    public static function createFromDateTime(DateTimeInterface $date): self
    {
        return new static($date->format(\DateTime::ATOM));
    }

    /**
     * Add days.
     *
     * @param int $days
     *
     * @return Date
     *
     * @throws \Exception
     */
    public function addDays(int $days): self
    {
        return $this->add(new DateInterval('P'.$days.'D'));
    }

    /**
     * Add hours.
     *
     * @param int $hours
     *
     * @return Date
     *
     * @throws \Exception
     */
    public function addHours(int $hours): self
    {
        return $this->add(new DateInterval('PT'.$hours.'H'));
    }

    /**
     * Get hours.
     *
     * @param int $hours
     *
     * @return Date
     *
     * @throws \Exception
     */
    public function getHours(): float
    {
        return (float) $this->format('H');
    }

    /**
     * Compute hours since another date.
     *
     * @param Date $date
     *
     * @return float
     */
    public function hoursSince(self $date): float
    {
        $interval = $this->diff($date);
        $hours = (float) $interval->h;
        if ($interval->i > 0) {
            $hours += (float) $interval->i / 60;
        }

        return $hours;
    }

    public function equals(self $other)
    {
        return $this === $other;
    }
}
