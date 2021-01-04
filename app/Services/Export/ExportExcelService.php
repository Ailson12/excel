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
     * Apenas os indices cabecalho e dados são obrigatorios
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
            array_unshift($data['dados'], $data['cabecalho']);
            $nomeArquivo = $data['nome-arquivo'] ?? 'document.xlsx';
            $iniciaCelula = $data['inicia-celula'] ?? 'A1';

            $spreadsheet = new Spreadsheet();

            $spreadsheet->getActiveSheet()
                ->fromArray(
                    $data['dados'],  
                    NULL,
                    $iniciaCelula
                );

            $coluna = 'A';
            for ($i = 0; $i < count($data['cabecalho']); $i++) {
                $spreadsheet->getActiveSheet()->getColumnDimension($coluna++)->setAutoSize(true);
            }

            if (isset($data['filtro'])) {
                $spreadsheet->getActiveSheet()->setAutoFilter($data['filtro']);
            }

            $estilos = $data['estilos'] ?? false;
            if ($estilos) {
                foreach ($estilos as $key => $value) {
                    $spreadsheet->getActiveSheet()->getStyle($key)->applyFromArray($value);
                }
            }
            
            $writer = new Xlsx($spreadsheet);

            header('Content-Type: application/vnd.ms-excel');
            header("Content-Disposition: attachment; filename=$nomeArquivo");
            $writer->save("php://output");
        } catch (Throwable $th) {
            LogService::generateLogError($th);
            throw new Exception('Erro ao gerar excel. Tente novamente', 500);
        }
    }
}