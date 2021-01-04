<?php

namespace App\Services;

use Throwable;
use Illuminate\Support\Facades\Log;

class LogService
{
   public static function generateLogError(Throwable $throwable)
   {
       Log::error([
           'message' => $throwable->getMessage(),
           'file' => $throwable->getFile(),
           'line' => $throwable->getLine()
       ]);
   }
}