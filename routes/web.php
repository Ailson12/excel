<?php

use Illuminate\Support\Facades\Route;

Route::get('/', function () {
    return view('welcome');
});

Route::get('excel', 'ExcelController@index');

Route::get('export-alunos-excel', 'AlunoController@exportAlunoExcel')->name('export.alunos.excel');