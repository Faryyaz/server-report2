<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use Carbon\Carbon;

/**
 * Generate the excel Report
 */
class Report extends Template {

    private $writer;
    private $serverData;

    public function __construct()
    {
        parent::__construct();

        $server = new ServerDataFormatter();
        $this->serverData = $server->get();

        $this->writer = new Xlsx($this->spreadsheet);
    }

    /**
     * Generate the excel file
     *
     * @return void
     */
    public function generate() : void
    {
        $today = Carbon::now()->format('d_m_Y');
        $this->setCellData();
        $this->processCellData();
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="Meteo_' . $today . '.xlsx"');
        $this->writer->save('php://output');
        exit();
    }

    /**
     * Dynamically Add data to the array cellData based on day's count
     *
     * @return array
     */
    protected function setCellData() : array
    {
        parent::setCellData();

        $days = $this->period->getDays();
        $count = 0;
        

        $serverAvaPos = [6, 15, 19, 23, 27, 31, 35];

        foreach ($days as $date => $day) {

            $countPos = 0;

            foreach($this->serverData as $datum) {

                $status_color = $this->getAvaStatusColor($datum['availability'][$date]->SLA_AVAILABILITY);
                $text = '';
                $fontName = 'Wingdings';
                if ($status_color === self::GREEN_COLOR) {
                    $text = 'J';
                } else {
                    $text = 'L';
                }

                array_push($this->cellData[$this->cellMap[$count]]['data'], [
                    $text, $serverAvaPos[$countPos], '', 'style' => 
                    $this->getStyle(
                        'HORIZONTAL_CENTER', 
                        self::BORDER_THIN, 
                        $status_color,
                        false,
                        self::BLACK_COLOR,
                        $fontName
                    )
                ]);
                $countPos++;
            }

            $this->setBatchAigCellData('batchAigLocation', $count, 8, $date);
            $this->setBatchAigCellData('batchAigDepartment', $count, 9, $date);
            $this->setBatchAigCellData('batchAigUsers', $count, 10, $date);

            $count++;
        }

        return $this->cellData;
    }

    /**
     * Set aig data in the given cell number
     *
     * @param string $dataSource
     * @param integer $cellMap
     * @param integer $cellPos
     * @param string $date
     * @return void
     */
    private function setBatchAigCellData(string $dataSource, int $cellMap, int $cellPos, string $date)
    {
        $dataSourceResult = $this->serverData['prod'][$dataSource];

        if (!empty($dataSourceResult)) {
            $completedDate = $dataSourceResult[$date];
            $style = $this->getStyle(
                'HORIZONTAL_CENTER', 
                self::BORDER_THIN,
                null,
                false,
                self::BLACK_COLOR
            );

            array_push($this->cellData[$this->cellMap[$cellMap]]['data'], [
                $completedDate, $cellPos, '', 
                'style' => $style
            ]);

        } else {
            $completedDate = '';
            array_push($this->cellData[$this->cellMap[$cellMap + 1]]['data'], [
                'Pas de rapport d\'import', $cellPos, '',
                'style' => [
                    'font' => [
                        'bold' => false,
                        'color' => ['rgb'=>self::BLACK_COLOR]
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_LEFT,
                    ],
                ]
            ]);

            $style = $this->getStyle(
                'HORIZONTAL_CENTER', 
                self::BORDER_THIN, 
                self::RED_COLOR
            );

            array_push($this->cellData[$this->cellMap[$cellMap]]['data'], [
                $completedDate, $cellPos, '', 
                'style' => $style
            ]);

        }
        
    }

}