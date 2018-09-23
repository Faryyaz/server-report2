<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

/**
 * Generate the excel Template
 */
class Template {

    protected $activeSheet;
    protected $period;
    protected $spreadsheet;
    
    const GREEN_COLOR = '339966';
    const RED_COLOR = 'ff0000';
    const BLUE_COLOR = '3366ff';
    const BLACK_COLOR = '000000';
    const YELLOW_COLOR = 'ffc000';
    const GREY_COLOR = 'a6a6a6';
    const BORDER_THIN = [
        'top' => 'BORDER_THIN', 
        'left' => 'BORDER_THIN', 
        'right' => 'BORDER_THIN',
        'bottom' => 'BORDER_THIN'
    ];
    const TITLE_STYLE_LEFT_ALIGNED = [ 
        'font' => [
            'bold' => true,
            'color' => ['rgb'=>'3366ff']
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_LEFT,
        ],
        'borders' => [
            'top' => [
                'borderStyle' => Border::BORDER_MEDIUM,
            ],
            'left' => [
                'borderStyle' => Border::BORDER_MEDIUM,
            ],
            'bottom' => [
                'borderStyle' => Border::BORDER_MEDIUM,
            ],
            'right' => [
                'borderStyle' => Border::BORDER_MEDIUM,
            ]
        ]
    ];   

    protected $cellData = [
        'A' => [
            'data' => [
                ['Bonjour,', 1, ''], // position 1( i.e. cell A1), no cell merging
                ['Veuillez trouver ci-dessous la météo des services Total ITSM NEXT, statut à 09h.', 2, 'E'], // position 2 (A2),  merge cells A to E
                ['Environnement PRODUCTION', 5, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Environnement PRE-PRODUCTION ', 14, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Environnement INTEGRATION', 18, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Environnement RECETTE', 22, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Environnement DEVELOPPEMENT', 26, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Environnement RFM', 30, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Environnement BAC-A-SABLE', 34, 'B',
                    'style' => self::TITLE_STYLE_LEFT_ALIGNED
                ],
                ['Numéro', 39, '','style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_THIN], 'right'=>['borderStyle' =>Border::BORDER_THIN], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 11, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 16, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 20, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 24, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 28, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 32, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 36, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 40, '','style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_THIN], 'right'=>['borderStyle' =>Border::BORDER_THIN], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 41, '','style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_THIN], 'right'=>['borderStyle' =>Border::BORDER_THIN], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 42, '','style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_THIN], 'right'=>['borderStyle' =>Border::BORDER_THIN], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 43, '','style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM], 'right'=>['borderStyle' =>Border::BORDER_THIN], 'left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 6, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 7, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 8, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 9, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 10, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 15, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 19, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 23, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 27, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 31, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['', 35, '','style'=>['borders'=>['left'=>['borderStyle' =>Border::BORDER_MEDIUM]]]]
            ],
            'width' => 'auto'
        ],
        'B' => [
            'data' => [
                ['Disponibilité de service', 6, ''],
                ['Sauvegarde', 7, ''],
                ['Batch AIG Location', 8, ''],
                ['Batch AIG Department', 9, ''],
                ['Batch AIG Users', 10, ''],
                ['TGS HYPERVISEUR', 11, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['Disponibilité de service', 6, ''],
                ['Sauvegarde', 7, ''],
                ['Disponibilité de service', 15, ''],
                ['Sauvegarde', 16, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['Disponibilité de service', 19, ''],
                ['Sauvegarde', 20, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['Disponibilité de service', 23, ''],
                ['Sauvegarde', 24, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['Disponibilité de service', 27, ''],
                ['Sauvegarde', 28, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['Disponibilité de service', 31, ''],
                ['Sauvegarde', 32, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]],
                ['Disponibilité de service', 35, ''],
                ['Sauvegarde', 36, '', 'style'=>['borders'=>['bottom'=>['borderStyle' =>Border::BORDER_MEDIUM]]]]                
            ],
            'width' => '22'
        ],
        'I' => [
            'data'=> [
                ['', 2, '', 'style'=>[
                        'borders'=>[
                            'bottom'=>['borderStyle' =>Border::BORDER_THIN],
                            'top'=>['borderStyle' =>Border::BORDER_THIN],
                            'left'=>['borderStyle' =>Border::BORDER_THIN],
                            'right'=>['borderStyle' =>Border::BORDER_THIN]
                        ],
                        'fill' => [
                            'fillType' => Fill::FILL_SOLID,
                            'color' => [
                                'rgb' => self::GREY_COLOR,
                            ]
                        ],
                    ]
                ],
                ['J', 3, '', 'style'=>[
                        'font' => [
                            'bold' => false,
                            'color' => ['rgb'=>self::BLACK_COLOR],
                            'name' => 'Wingdings'
                        ],
                        'alignment' => [
                            'horizontal' => Alignment::HORIZONTAL_CENTER,
                        ],
                        'borders'=>[
                            'bottom'=>['borderStyle' =>Border::BORDER_THIN],
                            'top'=>['borderStyle' =>Border::BORDER_THIN],
                            'left'=>['borderStyle' =>Border::BORDER_THIN],
                            'right'=>['borderStyle' =>Border::BORDER_THIN]
                        ],
                        'fill' => [
                            'fillType' => Fill::FILL_SOLID,
                            'color' => [
                                'rgb' => self::GREEN_COLOR,
                            ]
                        ],
                    ]
                ],
                ['K', 4, '', 'style'=>[
                        'font' => [
                            'bold' => false,
                            'color' => ['rgb'=>self::BLACK_COLOR],
                            'name' => 'Wingdings'
                        ],
                        'alignment' => [
                            'horizontal' => Alignment::HORIZONTAL_CENTER,
                        ],
                        'borders'=>[
                            'bottom'=>['borderStyle' =>Border::BORDER_THIN],
                            'top'=>['borderStyle' =>Border::BORDER_THIN],
                            'left'=>['borderStyle' =>Border::BORDER_THIN],
                            'right'=>['borderStyle' =>Border::BORDER_THIN]
                        ],
                        'fill' => [
                            'fillType' => Fill::FILL_SOLID,
                            'color' => [
                                'rgb' => self::YELLOW_COLOR,
                            ]
                        ],
                    ]
                ],
                ['L', 5, '', 'style'=>[
                        'font' => [
                            'bold' => false,
                            'color' => ['rgb'=>self::BLACK_COLOR],
                            'name' => 'Wingdings'
                        ],
                        'alignment' => [
                            'horizontal' => Alignment::HORIZONTAL_CENTER,
                        ],
                        'borders'=>[
                            'bottom'=>['borderStyle' =>Border::BORDER_THIN],
                            'top'=>['borderStyle' =>Border::BORDER_THIN],
                            'left'=>['borderStyle' =>Border::BORDER_THIN],
                            'right'=>['borderStyle' =>Border::BORDER_THIN]
                        ],
                        'fill' => [
                            'fillType' => Fill::FILL_SOLID,
                            'color' => [
                                'rgb' => self::RED_COLOR,
                            ]
                        ],
                    ]
                ],
                ['N/A', 6, '']
            ]
        ]
    ];

    /**
     * Excel columns, use for dynamic day
     */
    protected $cellMap = [
        'C', 'D', 'E', 'F', 'G', 'H'
    ];

    protected function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->activeSheet = $this->spreadsheet->getActiveSheet();
        $this->period = Period::getInstance();
    }


    /**
     * set values,style and size of excel sheet's cells
     *
     * @return void
     */
    protected function processCellData() : void
    {
        
        foreach ($this->cellData as $columnValue=>$cellValues) {

            if (isset($cellValues['width'])) {
                if ($cellValues['width'] === 'auto') {
                    $this->activeSheet->getColumnDimension($columnValue)->setAutoSize(true);
                } else {
                    $this->activeSheet->getColumnDimension($columnValue)->setWidth($cellValues['width']);
                }
            }

            foreach ($cellValues['data'] as $key=>$data) {
                
                $this->activeSheet->setCellValue($columnValue . $data[1], $data[0]);

                if ($data[2] !== '') {
                    $this->activeSheet->mergeCells($columnValue . $data[1] . ':' . $data[2] . $data[1]);
                }

                if (array_key_exists('style', $data)) {
                    if ($data[2] !== '') {
                        $this->activeSheet->getStyle($columnValue . $data[1] . ':' . $data[2] . $data[1])->applyFromArray($data['style']);
                    } else {
                        $this->activeSheet->getStyle($columnValue . $data[1])->applyFromArray($data['style']);
                    }
                }

            }
        }
    }


    /**
     * Dynamically Add data to the array cellData based on day's count
     *
     * @return array
     */
    protected function setCellData() : array
    {
        $days = $this->period->getDays();
        $count = 0;

        $endingCell = 'D';

        foreach ($days as $date => $day) {

            $this->cellData[$this->cellMap[$count]] = [
                'data' => [
                    ['Statut', 5, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['Statut', 14, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['Statut', 18, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['Statut', 22, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['Statut', 26, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['Statut', 30, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['Statut', 34, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ],
                    ['', 6, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 7, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 8, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 9, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 10, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 11, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 16, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 20, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 24, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 28, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 32, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                    ['', 36, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', self::BORDER_THIN)],
                ],
                'width' => '22'
            ];

            $this->cellData[$this->cellMap[$count + 1]] = [
                'data' => [
                    ['Notes', 5, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['Notes', 14, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['Notes', 18, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['Notes', 22, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['Notes', 26, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['Notes', 30, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['Notes', 34, '',
                        'style' => $this->getStyle('HORIZONTAL_LEFT')
                    ],
                    ['', 6, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 7, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 8, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 9, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 10, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 11, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])],
                    ['', 15, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 16, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])],
                    ['', 19, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 20, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])],
                    ['', 23, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 24, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])],
                    ['', 27, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 28, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])],
                    ['', 31, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 32, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])],
                    ['', 35, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM'])],
                    ['', 36, '', 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])]
                ],
                'width' => '50'
            ];

            if (count($days) > 1) {
                array_unshift(
                    $this->cellData[$this->cellMap[$count]]['data'],
                    [strtoupper($day), 4, '',
                        'style' => $this->getStyle('HORIZONTAL_CENTER')
                    ]
                );
            }

            $endingCell = $this->cellMap[$count + 1];

            $count++;
        }

        $incidentTableData = [
            ['Détails', 39, $endingCell, 'style' => $this->getStyle('HORIZONTAL_LEFT', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_THIN'], null, false, self::BLACK_COLOR)],
            ['', 40, $endingCell, 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_THIN'])],
            ['', 41, $endingCell, 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_THIN'])],
            ['', 42, $endingCell, 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_THIN'])],
            ['', 43, $endingCell, 'style' => $this->getStyle('HORIZONTAL_CENTER', ['right'=>'BORDER_MEDIUM', 'bottom'=>'BORDER_MEDIUM'])]
        ];

        $incidentTableHeader = ['Incidents P1 en cours - Aucun', 38, $endingCell,
            'style' => self::TITLE_STYLE_LEFT_ALIGNED
        ];

        if (isset($this->cellData['B'])) {
            foreach($incidentTableData as $tableData) {
                array_push($this->cellData['B']['data'], $tableData);
            }
        }

        if (isset($this->cellData['A'])) {
            array_push($this->cellData['A']['data'], $incidentTableHeader);
        }

        return $this->cellData;
    }

    /**
     * Return the style for text
     *
     * @param string $textAlignment HORIZONTAL_CENTER | HORIZONTAL_LEFT | HORIZONTAL_RIGHT
     * @param array $border ['top'=> 'BORDER_THICK', 'left'=> 'BORDER_THIN', 'right'=> 'BORDER_MEDIUM']
     * @param string $fill color to fill cell
     * @param array $fontBold true | false
     * @param array $fontColor 
     * @return array
     */
    protected function getStyle(
        string $textAlignment = 'HORIZONTAL_CENTER', 
        array $border = [
                            'top' => 'BORDER_MEDIUM', 
                            'left' => 'BORDER_MEDIUM', 
                            'right' => 'BORDER_MEDIUM',
                            'bottom' => 'BORDER_MEDIUM'
        ],
        string $fill = null,
        bool $fontBold = true,
        string $fontColor = self::BLUE_COLOR,
        string $fontName = 'Calibri'
    ) : array
    {
        return [
            'font' => [
                'bold' => $fontBold,
                'color' => ['rgb'=>$fontColor],
                'name' => $fontName
            ],
            'alignment' => [
                'horizontal' => constant('\PhpOffice\PhpSpreadsheet\Style\Alignment::' . $textAlignment),
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => isset($border['top']) ? constant('\PhpOffice\PhpSpreadsheet\Style\Border::' . $border['top']) : Border::BORDER_NONE,
                ],
                'left' => [
                    'borderStyle' => isset($border['left']) ? constant('\PhpOffice\PhpSpreadsheet\Style\Border::' . $border['left']) : Border::BORDER_NONE,
                ],
                'bottom' => [
                    'borderStyle' => isset($border['bottom']) ? constant('\PhpOffice\PhpSpreadsheet\Style\Border::' . $border['bottom']) : Border::BORDER_NONE,
                ],
                'right' => [
                    'borderStyle' => isset($border['right']) ? constant('\PhpOffice\PhpSpreadsheet\Style\Border::' . $border['right']) : Border::BORDER_NONE,
                ]
            ],
            'fill' => [
                'fillType' => $fill === null ? Fill::FILL_NONE : Fill::FILL_SOLID,
                'color' => [
                    'rgb' => $fill,
                ]
            ],
        ];
    }

    /**
     * Return the color based on availabilityValue, green (success) if more than threshold value else red (alert).
     *
     * @param integer $availabilityValue
     * @return string
     */
    protected function getAvaStatusColor(int $availabilityValue) : string
    {
        if ( $availabilityValue >= env('IPLABEL_SLA_AVAILABILITY_THRESHOLD') ) {
            return self::GREEN_COLOR; // green color
        }
        return self::RED_COLOR; // red color
    }


}