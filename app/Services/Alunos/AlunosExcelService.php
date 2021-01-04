<?php

namespace App\Services\Alunos;

use Exception;
use Throwable;
use App\Models\User;
use App\Services\LogService;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use App\Services\Export\ExportExcelService;

class AlunosExcelService
{
    public static function alunosExcel()
    {
        try {
            $arrayData = User::get(['name', 'email', 'created_at'])->take(10)->toArray();
            $totalLinhas = count($arrayData) + 1;

            $estilos = [
                "A1:C$totalLinhas" => [
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => '00000000']
                        ],
                        'fill' => [
                            'fillType' =>  Fill::FILL_SOLID,
                            'startColor' => ['argb' => 'FF4F81BD']
                        ]
                    ]
                ],
                "A1:C1" => [
                    'alignment' => [
                        'horizontal' => 'center'
                    ],
                    'font' => [
                        'bold' => true,
                        'color' => [
                            'argb' => Color::COLOR_WHITE
                        ]
                    ],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                            'argb' => 'FF4F81BD',
                        ]
                    ],
                ]
            ];

            // $data = [
            //     'nome-arquivo' => 'alunos.xlsx',
            //     'cabecalho' => ['Nome', 'Email ', 'Data de Criação'],
            //     'dados' => $arrayData,
            //     'estilos' => $estilos,
            //     'filtro' => 'A1:C1'
            // ];
            $data = [
                'cabecalho' => ['Nome', 'Email ', 'Data de Criação'],
                'dados' => [
                    [
                        [
                            'ailson', 'ailson@gmail.com', '2020-12-31'
                        ]
                    ],
                    [
                        [
                            'roberto', 'roberto@gmail.com', '2020-12-20'
                        ]
                    ]
                ],
            ];
            return ExportExcelService::exportExcel($data);
        } catch (Throwable $th) {
            LogService::generateLogError($th);
            throw new Exception($th->getMessage(), $th->getCode());
        }
    }
}