<?php

include_once "ze-reportes.php";
include_once "styles.php";

class excel_reportes implements ze_reportes_interface{    
    const LEYENDA_DERECHA = 1;
    const LEYENDA_DEBAJO = 2;
    const GRAFICO_DERECHA = 1;
    const GRAFICO_DEBAJO = 2;
    const HORIZONTAL = 'horizontal';
    const VERTICAL = 'vertical';
    const BARRA = 'barra';
    const LINEA = 'lineas';
    const AREA  = 'area';
    const TORTA = 'torta';
    
    private $encoding;
    private $titulo;
    private $subtitulo;
    private $columnas = array();   
    private $columnasMerge = array();
    private $objPHPExcel;
    private $sheet;
    private $numSheet;    
    private $error;
    private $fila_actual; 
    private $columna_actual;
    private $fila_encabezado;
    private $nombre_archivo;
    private $debug;
    private $html;
    private $alto_fila;
    private $alto_encabezados;
    private $area = array();
    private $orientacion = array();
    private $grafico;
    
    private $celda = array(
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
        "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV",
        "AW", "AX", "AY", "AZ"
    );
    
    private $alineacion = array(
        "centro" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        "derecha" => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        "izquierda" => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
    );
            
    private $formato = array(
        "entero" => PHPExcel_Style_NumberFormat::FORMAT_NUMBER,
        "decimal" => PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1,
        "moneda" => PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE,
        "fecha" => PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY                
    );
     
    public function __construct($encoding = "UTF8", $propiedades = array()){
        try {            
            $this->titulo = "Utilice el m�todo setTitulo(\$titulo)";
            $this->subtitulo = "";            
            $this->error = null;              
            $this->debug = false;   
            $this->html = "";
            $this->numSheet = 0;
            $this->encoding = $encoding;            
            $this->fila_actual = $this->fila_encabezado = 1;
            $this->orientacion[self::HORIZONTAL] = PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE;
            $this->orientacion[self::VERTICAL] = PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT;
            $this->grafico = false;
            $this->objPHPExcel = new PHPExcel();
            $this->objPHPExcel->setActiveSheetIndex($this->numSheet);
            $this->sheet = $this->objPHPExcel->getActiveSheet();                        
            $this->setDefaults();
            $this->setPropiedades($propiedades);
        } catch (Exception $ex) {
            $this->error = $ex->getMessage();
            $this->debug("Linea: ".__LINE__ . ": " . $ex->getMessage(), true);
            exit;
        }
    }
    
    
    private function debug($message){
        $html = <<<HTML
            <div style="border: 1px solid;
                margin: 10px 0px;
                padding:15px 10px 15px 50px;
                background-repeat: no-repeat;
                background-position: 10px center; 
                color: #D8000C; 
                background-color: #FFBABA; 
                width: 500px;
                position: absolute;
                left: 35%;
                background-image: url('/reporte/articulo3/ze-reportes/stop.png');">{$message}</div>
HTML;
        echo $html;                    
    }
    
    private function _echo($texto, $linea){
        if($this->debug == true){
            debug("L�nea {$linea}: ".$texto, true);
        }
    }
    
    public function modeDebug(){
        $this->debug = true;
    }
    
    public function nuevaHoja($hoja=0, $titulo=null){
        $this->numSheet = $hoja;
        $this->objPHPExcel->setActiveSheetIndex($this->numSheet);
        $this->sheet = $this->objPHPExcel->getActiveSheet();   
        $this->objPHPExcel->createSheet();
        $this->sheet->setTitle($titulo);
        $this->fila_actual = 1;
        $this->columnas = array();
    }
    
    
    public function setDefaults($alto = 15, $alto_encabezados = 20, $fuente = 'Arial', $tamano = 10){
        $this->alto_fila = $alto;
        $this->alto_encabezados = $alto_encabezados;
        $this->objPHPExcel->getDefaultStyle()->getFont()->setName($fuente);
        $this->objPHPExcel->getDefaultStyle()->getFont()->setSize($tamano);        
    }
    
    private function setPropiedades($propiedades){        
        if(isset($propiedades["nombre_archivo"])){
            $this->nombre_archivo = $propiedades["nombre_archivo"];
        }
        else{
            $this->nombre_archivo = "ze-reporte_excel.xlsx";
        }
        $creado_por = (isset($propiedades["creado_por"]))? utf8_encode($propiedades["creado_por"]): "Zecovery";
        $modificado_por = (isset($propiedades["modificado_por"]))? utf8_encode($propiedades["modificado_por"]): "";
        $asunto = (isset($propiedades["asunto"]))? utf8_encode($propiedades["asunto"]): "";
        $titulo = (isset($propiedades["titulo"]))? utf8_encode($propiedades["titulo"]): $this->titulo;
        $descripcion = (isset($propiedades["descripcion"]))? utf8_encode($propiedades["descripcion"]): "";
        $claves = (isset($propiedades["claves"]))? utf8_encode($propiedades["claves"]): "";
        $categoria = (isset($propiedades["categoria"]))? utf8_encode($propiedades["categoria"]): "";
        
        $this->objPHPExcel->getProperties()
            ->setCreator($creado_por)
            ->setLastModifiedBy($modificado_por)
            ->setTitle($titulo)
            ->setSubject($asunto)
            ->setDescription($descripcion)
            ->setKeywords($claves)
            ->setCategory($categoria);
        
        if($this->debug){
            $properties = array();
            $properties['Creator'] = $creado_por;
            $properties['LastModifiedBy'] = $modificado_por;
            $properties['Title'] = $titulo;
            $properties['Subject'] = $asunto;
            $properties['Description'] = $descripcion;
            $properties['Keywords'] = $claves;
            $properties['Category'] = $categoria;
            $properties['NOTA'] = "Las propiedades deben estar en codificaci�n UTF8, utilice utf8_encode";
            $this->debug($properties, true);
        }
    }
    
    public function setOrientacion($orientacion){
        if($orientacion== self::HORIZONTAL){
            $this->sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
        }
        else{
            $this->sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        }
    }
    
    public function setTamanoHoja($tamano){
        $this->sheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
    }
    
    public function setMargenes($arriba, $derecha, $izquierda, $abajo, $centerX=true, $centerY=false){
        $this->sheet->getPageMargins()->setTop($arriba);
        $this->sheet->getPageMargins()->setRight($derecha);
        $this->sheet->getPageMargins()->setLeft($izquierda);
        $this->sheet->getPageMargins()->setBottom($abajo);
        $this->sheet->getPageSetup()->setHorizontalCentered($centerX);
        $this->sheet->getPageSetup()->setVerticalCentered($centerY);
    }
            
    public function printHeaderFooter($header, $footer){
        $this->sheet->getHeaderFooter()->setOddHeader($header);
        $this->sheet->getHeaderFooter()->setOddFooter($footer);
    }
    
    public function setAreaImpresion($desde, $hasta){
        $this->sheet->getPageSetup()->setPrintArea("{$desde}:{$hasta}");
    }
    
    private function getEstilo($estilo){
        global $_style;        
        return($_style[$estilo]);            
    }
            
    public function setEncoding($encoding){
        $this->encoding = strtolower($encoding);
    }
    
    public function setTitulo($titulo){        
        $titulo = ($this->encoding  != "utf8")? utf8_encode($titulo): $titulo;
        $celda = "A" . $this->fila_actual;        
        $this->sheet->setCellValue($celda, $titulo);
        $this->sheet->getStyle($celda)->applyFromArray($this->getEstilo('titulo'));
        
        $encabezados = $this->columnas;
        $cols = count($encabezados);
        $this->sheet->mergeCells($celda.':'.$this->celda[$cols-1].$this->fila_actual);        
        $this->fila_actual++;
        
        $this->html .= "<h1>{$titulo}</h1>";
    }
    
    public function setSubTitulo($subtitulo){
        $subtitulo = ($this->encoding  != "utf8")? utf8_encode($subtitulo): $subtitulo;
        $celda = "A" . $this->fila_actual;
        $this->sheet->setCellValue($celda, $subtitulo);
        $this->sheet->getStyle($celda)->applyFromArray($this->getEstilo('subtitulo'));
        
        $encabezados = $this->columnas;
        $cols = count($encabezados);
        $this->sheet->mergeCells($celda.':'.$this->celda[$cols-1].$this->fila_actual);        
        $this->fila_actual++;
        
        $this->html .= "<h2>{$subtitulo}</h2>";
    }
    
    public function setEncabezados(){
        try {
            $encabezados = $this->columnas;
            $cols = count($encabezados);

            if ($cols == 0) {
                throw new Exception("No se ha configurado los encabezados");
            }

            $row = $this->fila_actual+1;
            $col = 0;
            $celda = $this->celda[$col] . $row;
            $this->fila_encabezado = $row;
            $this->html .= "<thead><tr>";
            foreach($encabezados as $encabezado){
                if(!is_numeric($encabezado["ancho"])){
                    throw new Exception("Ancho {$encabezado['ancho']} no valido para la columna {$encabezado['nombre']}");
                }
                
                $texto = ($this->encoding  != "utf8")?  utf8_encode($encabezado["nombre"]):  $encabezado["nombre"];                
                $this->sheet->setCellValue($this->celda[$col] . $row, $texto);
                
                $this->sheet->getStyle($this->celda[$col])->getAlignment()->setWrapText(true);
                $this->sheet->getColumnDimension($this->celda[$col])->setWidth($encabezado["ancho"]);
                
                $this->html .= "<td><small>".$this->celda[$col] . $row."</small><p><b>{$texto}</b></p></td>";
                $col++;
            }
            $this->html .= "</tr></thead>";
            
            $celda .= ":".$this->celda[$col - 1] . $row;
            $this->fila_actual = $row;
            $this->columna_actual = $col;            
            $this->sheet->getStyle($celda)->applyFromArray($this->getEstilo('encabezados'));
            return(true);
        }
        catch(Exception $ex){
            $this->error = $ex->getMessage();
            $this->debug($this->error, true);            
            exit;
        }
    }
    
    public function setEncabezadoMerge(){
        try {
            $encabezados = $this->columnasMerge;;
            $cols = count($encabezados);

            if ($cols == 0) {
                return;
            }

$enc=1;

            foreach($encabezados as $encabezado2){
			$row = $this->fila_actual+1;
            $col = 0;
            $celda = $this->celda[$col] . $row;
            foreach($encabezado2 as $encabezado){
                $texto = ($this->encoding  != "utf8")?  utf8_encode($encabezado["nombre"]):  $encabezado["nombre"];                
                $col_ini = $encabezado["desde"];
                $col_fin = $encabezado["hasta"];
                $celda = $this->celda[$col_ini].$row.":".$this->celda[$col_fin].$row;
				//print_r($encabezado);
				//print_r($this->celda[$col_ini]);
                $this->sheet->mergeCells($celda);
                $this->sheet->setCellValue($this->celda[$col_ini] . $row, $texto);
                $this->sheet->getStyle($celda)->getAlignment()->setHorizontal($this->alineacion[$encabezado["alineacion"]]);
                $this->sheet->getStyle($celda)->applyFromArray($this->getEstilo('encabezados_agrupados_'.$enc));
                $col++;
            }
            $this->sheet->getRowDimension($row)->setRowHeight($this->alto_encabezados);
            $this->fila_actual = $row; 
			$enc++;
			}			
            return(true);
        }
        catch(Exception $ex){
            $this->error = $ex->getMessage();
            $this->debug($this->error, true);            
            exit;
        }
    }
    
    public function setGrafico($propiedades){        
        try{
            $titulo =  ($propiedades["titulo"])? utf8_encode($propiedades["titulo"]) : "";        
            $axis = $propiedades["abscisa"];
            $columns = $propiedades["columnas"];
            $tamano = ($propiedades["dimensiones"])? $propiedades["dimensiones"]: array("alto" => 400, "ancho" => 400);        
            $posicion = (isset($propiedades['posicion']))? $propiedades['posicion']: 2;
            
            //Si el grafico va abajo, la fila donde comienza es la siguiente a la actual, si no es la fila donde se dibuj� el encabezado
            $fila = ($posicion == 2)? $this->fila_actual+1 : $this->fila_encabezado ;        
            //Si el grafico va al lado, la columna donde comienza es el numero de columnas de la data mas 1, sino es la primera columna
            $col = ($posicion == 1)? $this->numColumnas()+1: 0;
            $primera_fila = $fila;
            $primera_col = $col;
            
            $xal = array(); //abcisas
            
            for ($i = 0; $i < count($axis); $i++) {
                $celdas = $this->celda[$axis[$i]] . '$' . $this->area['fila_ini'] . ':$' . $this->celda[$axis[$i]] . '$' . $this->area['fila_fin'];
                $this->_echo($celdas, __LINE__);
                $xal[] = new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$' . $celdas, NULL, 90);                                
            }

            $dsl = array();  //nombre de las series
            $dsv = array();   //datos
            for ($i = 0; $i < count($columns); $i++) {
                $celda = $this->celda[$columns[$i]] . '$' . $this->fila_encabezado;                
                $this->_echo($celda, __LINE__);
                $dsl[] = new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$' . $celda, NULL, 1);  
                $celdas = $this->celda[$columns[$i]] . '$' . $this->area['fila_ini'] . ':$' . $this->celda[$columns[$i]] . '$' . $this->area['fila_fin'];
                $this->_echo($celdas, __LINE__);
                $dsv[] = new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$' . $this->celda[$columns[$i]] . '$' . $this->area['fila_ini'] . ':$' . $this->celda[$columns[$i]] . '$' . $this->area['fila_fin'], NULL, 90);                
            }


            switch($propiedades["tipo"]){
                case self::LINEA:
                    $ds = new \PHPExcel_Chart_DataSeries(\PHPExcel_Chart_DataSeries::TYPE_LINECHART,  \PHPExcel_Chart_DataSeries::GROUPING_STANDARD, range(0, count($dsv) - 1),  $dsl,  $xal,  $dsv);
                    break;
                case self::AREA:
                    $ds = new \PHPExcel_Chart_DataSeries(\PHPExcel_Chart_DataSeries::TYPE_AREACHART,  \PHPExcel_Chart_DataSeries::GROUPING_STANDARD, range(0, count($dsv) - 1),  $dsl,  $xal,  $dsv);
                    break;
                case self::TORTA:
                    $ds = new \PHPExcel_Chart_DataSeries(\PHPExcel_Chart_DataSeries::TYPE_PIECHART,  \PHPExcel_Chart_DataSeries::GROUPING_STANDARD, range(0, count($dsv) - 1),  $dsl,  $xal,  $dsv);
                    break;
                case self::BARRA:
                default: 
                    $ds = new \PHPExcel_Chart_DataSeries(\PHPExcel_Chart_DataSeries::TYPE_BARCHART,  \PHPExcel_Chart_DataSeries::GROUPING_STANDARD, range(0, count($dsv) - 1),  $dsl,  $xal,  $dsv);
            }

            $pa = new \PHPExcel_Chart_PlotArea(NULL, array($ds));
            $legend = new \PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
            $title = new \PHPExcel_Chart_Title($titulo);
            
            $axisX = (isset($propiedades['labelX']))? $propiedades['labelX']: "";
            $xAxisLabel =  new \PHPExcel_Chart_Title($axisX);
            
            $axisY = (isset($propiedades['labelY']))? $propiedades['labelY']: "";            
            $yAxisLabel = new \PHPExcel_Chart_Title($axisY);
                        
            $chart = new \PHPExcel_Chart(
                    'chart1', $title, $legend, $pa, true, 0, $xAxisLabel, $yAxisLabel
            );

            $posIni = $this->celda[$col].$fila;            
            $posFin = $this->celda[$col+$tamano["ancho"]].($tamano["alto"]+$fila);
            $this->_echo($posIni, __LINE__);
            $this->_echo($posFin, __LINE__);
            $chart->setTopLeftPosition($posIni);
            $chart->setBottomRightPosition($posFin);
            $this->sheet->addChart($chart);
            if($fila > $this->fila_actual){
                $this->fila_actual = $fila;
            }
            $this->grafico = true;

        }
        catch(Exception $ex){
            $this->error = $ex->getMessage();
            $this->debug($this->error, true);            
            exit;
        }

    }
    
    public function setColumnaMerge($columna){
            $this->columnasMerge[] = $columna;
    }
    
    public function setColumna($columna){
        $this->columnas[] = $columna;
    }
    
    private function numColumnas(){        
        return(count($this->data[0]));
    }
    
    
    public function setLeyenda($leyenda, $posicion=self::LEYENDA_DERECHA){ //1=Derecha, 2=Abajo        
        try{            
            if($leyenda == null || count($leyenda) == 0){
                return;
            }
            
            if($posicion != self::LEYENDA_DERECHA && $posicion != self::LEYENDA_DEBAJO){
                throw new Exception("Valor {$posicion} no v�lido para la posici�n, debe ser LEYENDA_DERECHA � LEYENDA_DEBAJO");
            }
            
            //Si la leyenda va abajo, la fila donde comienza es la siguiente a la actual, si no es la fila donde se dibuj� el encabezado
            $fila = ($posicion == 2)? $this->fila_actual+1 : $this->fila_encabezado ;        
            //Si la leyenda va al lado, la columna donde comienza es el numero de columnas de la data mas 1, sino es la primera columna
            $col = ($posicion == 1)? $this->numColumnas()+1: 0;
            $primera_fila = $fila;
            $primera_col = $col;
            
            //Muestra los encabezados de la Leyenda
            $html = "<table width='100%' border='1' cellpadding='5'><thead><tr>";                        
            foreach ($leyenda["encabezados"] as $encabezado) {
                $this->sheet->getColumnDimension($this->celda[$col])->setWidth($encabezado['ancho']);
                $texto = ($this->encoding != "utf8") ? utf8_encode($encabezado['nombre']) : $encabezado['nombre'];
                $celda = $this->celda[$col].$fila;
                $this->sheet->setCellValue($celda, $texto);
                $this->sheet->getStyle($celda)->getAlignment()->setHorizontal($this->alineacion["centro"]);
                $this->sheet->getRowDimension($fila)->setRowHeight($this->alto_fila);
                $html .= "<td><small>{$celda}</small><p><b>{$texto}</b></p></td>";
                $col++;
            }
            $fila++;            
            $html .= "</tr></thead><tbody>";            
            
            //Aplicar el estilo a las cabeceras
            $celdas = $this->celda[$primera_col] . $primera_fila . ":" . $this->celda[$col-1] . ($primera_fila);            
            $this->sheet->getStyle($celdas)->applyFromArray($this->getEstilo('encabezados_leyenda'));
            //Muestra la tabla de leyenda
            $html .= "<tbody>";
            foreach ($leyenda["datos"] as $datos) {                        
                $col = $primera_col;
                $html .= "<tr>";
                foreach ($datos as $texto) {                    
                    $texto = ($this->encoding != "utf8") ? utf8_encode($texto) : $texto;
                    $celda = $this->celda[$col].$fila;
                    $this->sheet->setCellValue($celda, $texto);
                    $this->sheet->getStyle($celda)->getAlignment()->setHorizontal($this->alineacion["izquierda"]);
                    $this->sheet->getRowDimension($fila)->setRowHeight($this->alto_fila);
                    $html .= "<td><small>{$celda}</small><p><b>{$texto}</b></p></td>";
                    $col++;
                }
                $fila++;
                $html .= "</tr>";
            }
            
            $html .= "</tbody></table><tr>";
            
            //Aplica el borde a la tabla
            $celdas = $this->celda[$primera_col] . ($primera_fila+1). ":" . $this->celda[$col-1] . ($fila-1);            
            $this->sheet->getStyle($celdas)->applyFromArray($this->getEstilo('cuerpo'));

            if($fila > $this->fila_actual){
                $this->fila_actual = $fila;
            }
            
            if($this->debug){
                echo $html;
            }
        }
        catch(Exception $ex){
            $this->error = $ex->getMessage();
            $this->debug(__LINE__.": ".$this->error, true);            
            exit;
        }
    }
    
    
    public function setReporte($datos, $opciones=array()){
        try {            
            if(!$datos || count($datos) == 0){
                throw new Exception("No hay datos para el reporte");
            }

            $this->data = $datos;
            $totales = array("TOTALES");
            $this->area = array();

            $this->setEncabezadoMerge(); //Lama a la funcion que muestra los encabezado agrupados
            $this->setEncabezados();  //Llama a mostrar los encabezados de las tablas
            $col_ini = 0;        //La columna inicial siempre sera 0
            $cols = count($this->columnas);   //Cantidad de columnas de la tabla

            $row_ini = $this->fila_actual + 1;   //Porque se colocan los titulo en la primera fila
            $this->fila_actual++; //Se mueve la fila actual
            $row = $row_ini; //Se asigna $row_ini a la primera fila


            $this->html .= "<tbody>";
            foreach ($datos as $data) {
                $col = $col_ini;
                $this->html ."<tr>";
                foreach ($data as $texto) {                    
                    $celda = $this->celda[$col] . $row;

                    if(isset($this->alineacion[$this->columnas[$col]["alineacion"]])){                    
                        $this->sheet->getStyle($celda)->getAlignment()->setHorizontal($this->alineacion[$this->columnas[$col]["alineacion"]]);
                    }
                    
                    //Si se configura un formato personalizado
                    if(isset($this->columnas[$col]["formato"])){
                        $this->sheet->getStyle($celda)->getNumberFormat()->setFormatCode($this->columnas[$col]["formato"]);
                    }
                    else  //En caso que se defina un formato de acuerdo al tipo (entero, decimal, fehca, moneda)
                    {
                        if(isset($this->columnas[$col]["tipo"]) && $this->columnas[$col]["tipo"] != "texto"){
                            $this->sheet->getStyle($celda)->getNumberFormat()->setFormatCode($this->formato[$this->columnas[$col]["tipo"]]);                            
                        }
                        else{
                            if($this->encoding != "utf8") {
                                $texto = utf8_encode($texto);
                            }
                        }
                    }
                    $this->sheet->setCellValue($celda, $texto);
                    $this->html .= "<td><small>{$celda}</small><p>{$texto}</p></td>";
                    $col++;                    
                }
                $this->sheet->getRowDimension($row)->setRowHeight($this->alto_fila);
                if($row%2 == 0){
                    $this->sheet->getStyle($this->celda[$col_ini].$row.":".$this->celda[$col-1].$row)->applyFromArray($this->getEstilo('fila_par'));
                }
                else{
                    $this->sheet->getStyle($this->celda[$col_ini].$row.":".$this->celda[$col-1].$row)->applyFromArray($this->getEstilo('fila_impar'));
                }
                $this->html .= "</tr>";
                $row++;                
            }            
            $this->html .= "</tbody>";
            
            $cols--;
            $celdas = $this->celda[$col_ini] . $row_ini . ":" . $this->celda[$cols] . ($row - 1);
            $this->sheet->getStyle($celdas)->applyFromArray($this->getEstilo('cuerpo'));            
            $this->_echo("Area de los datos ".$celdas, __LINE__);            
            $this->area['fila_fin'] = $row - 1;
            $this->fila_actual = $row+1;
            $this->columna_actual = $cols;
            $this->area = array("fila_ini" => $row_ini, "col_ini" => $this->celda[$col_ini], "fila_fin" => $row-1, "col_fin" => $this->celda[$cols]);
        } catch (Exception $e) {
            $this->error = "Linea ".__LINE__.": ".$e->getMessage();
            $this->debug($this->error, true);            
            return(false);
        }
    }
    
    /**
     * Funcion que crea una linea en blanco
     */
    public function lineaBlanco(){
        $this->fila_actual++;
    }
    
    public function showTabla(){
        echo "<table width='100%' border='1' cellpadding='5'>{$this->html}</table>";
    }
    
    public function show($tipo='xls'){
        $tipo = strtolower($tipo);
        
        try{
            if($this->error){                               
                throw new Exception($this->error);
            }

            if($this->debug){
                $this->showTabla();
                 throw new Exception("Est� habilitado el debug: Debe deshabilitarlo para poder mostrar el archivo {$tipo}");            
            }
        
            if (headers_sent()) {
                throw new Exception("Se ha enviado un texto antes de imprimir el archivo");
            }        

            if($tipo != 'xls' &&  $tipo != 'xlsx' && $tipo != 'pdf'){
                throw new Exception("Tipo de archivo {$tipo} no v�lido");
            }
            
            ob_clean();
            if(strtolower($tipo) == 'xls' || strtolower($tipo) == 'xlsx'){
                $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
                $objWriter->setIncludeCharts($this->grafico); 
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="'.$this->nombre_archivo.'"');
                header('Cache-Control: public, must-revalidate, max-age=0');
                $objWriter->save('php://output');
            }
            elseif(strtolower($tipo) == 'pdf'){
                $rendererName = PHPExcel_Settings::PDF_RENDERER_TCPDF;
                $rendererLibrary = 'tcpdf.php';
                $rendererLibraryPath = dirname($_SERVER['DOCUMENT_ROOT']) .'/include/base/tcpdf/'.$rendererLibrary;
                
                /*if (!PHPExcel_Settings::setPdfRenderer(	$rendererName,	$rendererLibraryPath)) {
                    $this->debug(
                        "NOTICE: Please set the $rendererName and $rendererLibraryPath values <br/>
                        at the top of this script as appropriate for your directory structure"
                    );
                }*/
                
                $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'PDF');
                header('Content-Type: application/pdf');                
                header('Content-Disposition: attachment;filename="'.$this->nombre_archivo.'"');
                header('Cache-Control: public, must-revalidate, max-age=0');
                $objWriter->save('php://output');
            }
        } catch (Exception $e) {
            $this->error = "Linea ".__LINE__.": ".$e->getMessage();
            $this->debug($this->error, true);                        
        }
    }
    
}

