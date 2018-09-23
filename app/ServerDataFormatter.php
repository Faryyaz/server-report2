<?php

namespace App;

use Carbon\Carbon;

/**
 * Get formatted data for report
 */
class ServerDataFormatter {

    private $data = [];

    /**
     * Return the formatted data from different servers
     *
     * @return array
     */
    public function get() : array
    {
        $period = Period::getInstance();
        $dates = $period->get();
        
        return $this->data = [
            'prod' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_PROD_MONITOR_ID'),
                    $dates
                ),
                'batchAigLocation' => $this->getBatchImportDates(
                    env('SERVICENOW_IMPORT_AIG_LOCATION_SYSID'), 
                    $dates
                ),
                'batchAigDepartment' => $this->getBatchImportDates(
                    env('SERVICENOW_IMPORT_AIG_DEPARTMENT_SYSID'), 
                    $dates
                ),
                'batchAigUsers' => $this->getBatchImportDates(
                    env('SERVICENOW_IMPORT_AIG_USERS_SYSID'), 
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
            'preprod' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_PPROD_MONITOR_ID'),
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
            'int' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_INT_MONITOR_ID'),
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
            'rec' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_REC_MONITOR_ID'),
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
            'dev' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_DEV_MONITOR_ID'),
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
            'form' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_FORM_MONITOR_ID'),
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
            'bas' => [
                'availability' => $this->getAvailability(
                    env('IPLABEL_BAS_MONITOR_ID'),
                    $dates
                ),
                'backup' => Server::getBackupStatus()
            ],
        ];
    }

    /**
     * Get the server availability based on input dates
     *
     * @param string $monitorId
     * @param array $dates
     * @return array
     */
    private function getAvailability(string $monitorId, array $dates) : array
    {
        $availability = [];
        for($i = 0; $i < count($dates) - 1; $i++) {

            $startDate = $dates[$i];
            $endDate = $dates[$i + 1];

            $availability[$endDate] = Server::getStatus(
                                $monitorId,
                                $startDate . ' 09:00:00',
                                $endDate . ' 09:00:00'
                            );

        }

        return $availability;
    }

    /**
     * Return the completed dates for batch imports
     *
     * @param string $sysId
     * @param array $dates
     * @return array
     */
    public function getBatchImportDates(string $sysId, array $dates) : array
    {
        $completedDates = [];

        for($i = 1; $i < count($dates); $i++) {
            $batchImportData = Server::getBatchImportData($sysId, $dates[$i]);
            if (!empty($batchImportData)) {
                $completedDate = Carbon::createFromFormat(
                    'Y-m-d H:i:s', 
                    $batchImportData[0]->completed
                );
                $completedDates[$dates[$i]] = $completedDate->format('d-m-Y H:i:s');
            }
        }
        return $completedDates;
    }
}