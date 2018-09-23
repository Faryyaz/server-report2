<?php

namespace App;

use Carbon\Carbon;

class Period {

    private static $instance = null;
    private $dates = [];

    private function __construct()
	{
		// Not necessary, but making this private will block all public instantiations
    }
    
    public static function getInstance()
	{
		if (!self::$instance)
		{
			self::$instance = new self();
        }
        return self::$instance;
	}

    /**
     * Return all the date(s) based on weekend and public holidays
     *
     * @return array
     */
    public function get() : array
    {
        $yesterday = Carbon::yesterday();
        $today = Carbon::now();

        if (!empty($this->dates)) {
            return $this->dates;
        }

        if ($today->isWeekday() && !$this->isPublicHoliday($today->format('d/m/Y'))) {

            switch ( $today->dayOfWeek ) {
                case Carbon::MONDAY:
                $this->dates = $this->getWeekendAndHolidayDates($today);
                break;

                case Carbon::TUESDAY:
                $this->dates = [
                        $yesterday->format('d/m/Y'),
                        $today->format('d/m/Y')
                    ];

                    if ($this->isPublicHoliday($yesterday->format('d/m/Y'))) {
                        $this->dates = $this->getWeekendAndHolidayDates($yesterday);
                    }

                break;

                default:
                $this->dates = [
                        $yesterday->format('d/m/Y'),
                        $today->format('d/m/Y')
                    ];

                    if ($this->isPublicHoliday($yesterday->format('d/m/Y'))) {
                        array_unshift($this->dates, $yesterday->subDays(1)->format('d/m/Y')); 
                    }

                break;
            }

        }

        return $this->dates;
    }

    /**
     * Return all dates after a weekend + public holiday if consecutive
     *
     * @param [string] $mondayDate
     * @return array
     */
    private function getWeekendAndHolidayDates($mondayDate) : array
    {
        $dates = [];
        $date = Carbon::createFromFormat('Y-m-d H:i:s', $mondayDate);
        $today = Carbon::now();
        $monday = $date->format('d/m/Y');
        $sunday = $date->subDays(1)->format('d/m/Y');
        $saturday = $date->subDays(1)->format('d/m/Y');
        $friday = $date->subDays(1)->format('d/m/Y');
        $thursday = $date->subDays(1)->format('d/m/Y');

        $dates = [
            $friday,
            $saturday,
            $sunday
        ];

        if ($this->isPublicHoliday($friday)) {
            array_unshift($dates, $thursday);
        } 
        if ($this->isPublicHoliday($monday)) {
            array_push($dates, $monday);
        }

        array_push($dates, $today->format('d/m/Y'));

        return $dates;
    }

    /**
     * check if given date is a public holiday
     *
     * @param string $date
     * @return boolean
     */
    public function isPublicHoliday(string $date) : bool
    {
        $formattedDate = Carbon::createFromFormat('d/m/Y', $date)->format('Y-m-d');
        if ( PublicHolidays::where('date', $formattedDate)->exists() ) {
            return true;
        }
        return false;
    }

    /**
     * Return all date + day as an array example ["12/09/2018" => "mercredi"]
     *
     * @return array
     */
    public function getDays() : array
    {        
        $days = [];
        $dates = $this->get();
        for ($i = 1; $i < count($dates); $i++) {
            $dateFormatted = Carbon::createFromFormat('d/m/Y', $dates[$i]);
            $days[$dates[$i]] = $dateFormatted->formatLocalized('%A');
        }
        return $days;
    }
}