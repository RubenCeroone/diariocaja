<?php
 
 namespace App\Http\Controllers;

 use App\Http\Controllers\Controller;
 use App\Models\DiarioCaja;
 use Maatwebsite\Excel\Facades\Excel;
 use App\Exports\DiarioCajaExport;
 
 class BotonExcelController extends Controller
 {
     public function invoke()
     {
         return view('botonexcel');
     }
 
     public function exportarExcel()
     {
         $datos = DiarioCaja::all();
         return Excel::download(new DiarioCajaExport($datos), 'DiarioCaja.xlsx');
     }
 }