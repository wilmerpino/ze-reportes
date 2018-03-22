<?php

interface ze_reportes_interface{
    public function setEncoding($encoding);
    
    public function setTitulo($titulo);
    
    public function setSubTitulo($titulo);
    
    public function setEncabezados();
    
    public function setGrafico($propiedades);
    
    public function setColumna($columna);
    
    public function setReporte($datos, $opciones=array());
    
    public function nuevaHoja($hoja=0,$titulo=null);
    
    public function show($tipo=null);
    
    public function setColumnaMerge($columna);
    
    public function setEncabezadoMerge();
}
