<?php

namespace Sede2\Excel;

use Maatwebsite\Excel\Facades\Excel;

class ExcelCreator{

    protected $arrayData;
    protected $excelName;
    protected $sheetName;
    protected $format;
    protected $view;

    private function __construct($arrayData, $view = '', $excelName = 'excel', $sheetName = 'sheet1', $format = 'xls'){
        $this->arrayData = $arrayData;
        $this->excelName = $excelName;
        $this->sheetName = $sheetName;
        $this->format = $format;
        $this->view = $view;
    }

    public static function create($arrayData, $view = '', $excelName = 'excel', $sheetName = 'sheet1', $format = 'xls'){
        $excel = new ExcelCreator($arrayData, $view, $excelName, $sheetName, $format);
        return $excel->createExcel();
    }

    private function createExcel(){
        if($this->view) return $this->createFromView();
        else if($this->arrayData) return $this->createFromArray();
    }

    private function createFromArray(){
        $array = $this->arrayData;
        $sheetName = $this->sheetName;
        return  Excel::create($this->excelName, function($excel) use ($array, $sheetName) {
            $excel->sheet($sheetName, function($sheet) use ($array) {
                $sheet->fromArray($array);
            });
        })->export($this->format);
    }

    private function createFromView(){
        $array = $this->arrayData;
        $sheetName = $this->sheetName;
        $view = $this->view;
        return Excel::create($this->excelName, function($excel) use ($array, $sheetName, $view) {
            $excel->sheet($sheetName, function($sheet) use ($array, $view) {
                $sheet->loadView($view, compact('array'));
            });
        })->export($this->format);
    }
}