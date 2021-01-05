<?php

namespace App\Services\Export;

use Exception;
use Throwable;
use App\Services\LogService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExportExcelService
{
    /**
     * ===========================================================
     * Adicionar Gráfico no Excel
     *  https://github.com/PHPOffice/PhpSpreadsheet/blob/master/samples/Chart/33_Chart_create_pie.php
     * ===========================================================
     * Apenas os indices cabecalho e dados são obrigatorios
     * 
     * ===========================================================
     * O indice inicia-celula é utilizado para definir apartir de qual celula o Excel deve ser montado
     * Ex: 'inicia-celula' => 'A3'
     * 
     * ===========================================================
     * O indice filtro deve ser definido em um intervalo de células
     * Ex: 'filtro' => 'A1:B2'
     * 
     * ===========================================================
     * O indice estilo deve ser um array tendo como chaves as celulas que receberão os estilos e como valor o proprio estilo.
     * Ex: 'estilos' => [
     *     'A1:B2' => [
     *           'borders' => [
     *              'allBorders' => [
     *                  'borderStyle' => Border::BORDER_THIN,
     *                  'color' => ['argb' => '00000000']
     *              ]
     *          ]
     *     ]
     * ]
     */

    public static function exportExcel(Array $data)
    {
        try {
            $nomeArquivo = $data['nome-arquivo'] ?? 'document.xlsx';
            $planilhas = $data['planilhas'];


            $spreadsheet = new Spreadsheet();

            foreach ($planilhas as $key => $planilha) {
                $iniciaCelula = $planilha['inicia-celula'] ?? 'A1';

                $sheet = "sheet_$key";

                if ($key == 0) {
                    $sheet = $spreadsheet->getActiveSheet();
                } else {
                    $sheet = $spreadsheet->createSheet();
                }

                array_unshift($planilha['dados'], $planilha['cabecalho']);

                $sheet->fromArray(
                    $planilha['dados'],
                    [NULL],
                    $iniciaCelula
                );

                if (isset($planilha['titulo'])) {
                    $sheet->setTitle($planilha['titulo']);
                }

                $coluna = $iniciaCelula[0];
                for ($i = 0; $i < count($planilha['cabecalho']); $i++) {
                    $sheet->getColumnDimension($coluna++)->setAutoSize(true);
                }

                if (isset($planilha['filtro'])) {
                    $sheet->setAutoFilter($planilha['filtro']);
                }
    
                $estilos = $planilha['estilos'] ?? false;
                if ($estilos) {
                    foreach ($estilos as $key => $value) {
                        $sheet->getStyle($key)->applyFromArray($value);
                    }
                }
            }
            
            return new Xlsx($spreadsheet);
        } catch (Throwable $th) {
            LogService::generateLogError($th);
            throw new Exception('Erro ao gerar excel. Tente novamente', 500);
        }
    }
}