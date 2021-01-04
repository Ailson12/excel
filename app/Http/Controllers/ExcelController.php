<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Throwable;

class ExcelController extends Controller
{
    public function index()
    {
        try {
          return ;
        } catch (Throwable $th) {
            return response()->json([$th->getMessage()], $th->getCode());
        }
    }
}
