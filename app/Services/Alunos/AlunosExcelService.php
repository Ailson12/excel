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
            $arrayData = User::get(['name', 'email', 'created_at'])->take(100)->toArray();

            $totalLinhas = count($arrayData) + 1;

            $estilosPlanilha1 = [
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

            $arrayData2 = [
                ['Status', 'Quantidade', 'Porcentagem'],
                ['ativo', 'roberto@gmail.com', '2020-11-31'],
                ['inativo', 'feitosa@gmail.com', '2020-10-31'],
                [NULL, NULL, NULL],
                ['Perfis', 'Quantidade', 'Porcentagem'],
                ['admin', 'roberto@gmail.com', '2020-11-31'],
                ['callcenter', 'feitosa@gmail.com', '2020-10-31'],
            ];

            $planilha1 = [
                'cabecalho' => ['Nome', 'Email', 'Data de Criação'],
                'titulo' => 'Usuários',
                'dados' => $arrayData,
                'estilos' => $estilosPlanilha1,
                'filtro' => 'A1:C1',
            ];

            $planilha2 = [
                'cabecalho' => ['Nome ', 'Email', 'Data de Criação'],
                'titulo' => 'Perfil',
                'dados' => $arrayData2,
                'tabelas' => [

                ],
                'filtro' => 'A1:C1'
            ];

            $data = [
                'nome-arquivo' => 'alunos.xlsx',
                'planilhas' => [
                    $planilha1,
                    $planilha2
                ]
            ];
            
            $exportExcel = ExportExcelService::exportExcel($data);

            header('Content-Type: application/vnd.ms-excel');
            header("Content-Disposition: attachment; filename=alunos.xlsx");
            $exportExcel->save("php://output");
        } catch (Throwable $th) {
            LogService::generateLogError($th);
            throw new Exception($th->getMessage(), $th->getCode());
        }
    }
}