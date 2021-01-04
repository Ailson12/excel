<?php

namespace App\Http\Controllers;

use Throwable;
use App\Services\Alunos\AlunosExcelService;

class AlunoController extends Controller
{
    public function exportAlunoExcel()
    {
        try {
            return AlunosExcelService::alunosExcel();
        } catch (Throwable $th) {
            return response()->json([$th->getMessage()], $th->getCode());
        }
    }
}
