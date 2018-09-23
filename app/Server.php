<?php

namespace App;

use Carbon\Carbon;

//Get the data from the different servers
class Server {

    /**
     * Return the status of the given $monitorId
     *
     * @param integer $monitorId
     * @param string $startDate
     * @param string $endDate
     * @return void
     */
    public static function getStatus(int $monitorId, string $startDate, string $endDate)
    {
        $url = env('IPLABEL_API_URL') . 'Get_KPI/';
        $username = env('IPLABEL_API_USERNAME');
        $password = env('IPLABEL_API_PASSWORD');
        $query = [
            'monitor_id' => $monitorId,
            'date_value1' => $startDate,
            'date_value2' => $endDate,
        ];

        $serverData = new ServerData(
            $url,
            $username,
            $password,
            $query
        );
        return $serverData->get()->Ipln_WS_REST_datametrie->Get_KPI->response;
    }

    public static function getBackupStatus()
    {
        // backup server logic here
        return [];
    }

    /**
     * return batchImports data from serviceNow
     *
     * @param string $sysId
     * @param string $date
     * @return void
     */
    public static function getBatchImportData(string $sysId, string $date)
    {
        $date = Carbon::createFromFormat('d/m/Y', $date)->format('Y-m-d');
        $url = env('SERVICENOW_API_URL') . '/now/table/sys_import_set_run';
        $username = env('SERVICENOW_API_USERNAME');
        $password = env('SERVICENOW_API_PASSWORD');
        $query = [
            'sysparm_query' => "set.data_source={$sysId}^completedLIKE{$date}",
            'sysparm_limit' => '1',
            'sysparm_display_value' => 'true'
        ];
        $serverData = new ServerData(
            $url,
            $username,
            $password,
            $query
        );
        return $serverData->get()->result;
    }
    
}