<?php

include_once "ze-reportes.php";

class html_reportes implements ze_reportes_interface{    
    private $encode;
    private $titulo;
    private $subtitulo;
    private $columnas = array();    
    private $data = array();
    private $error;    
    private $html;
    private $encoding;
    
    public function __construct($encoding = "utf8") {
        try {
            $this->titulo = "Utilice el método setTitulo(\$titulo)";
            $this->subtitulo = "";            
            $this->error = null;              
            $this->html = "";                
            $this->encode = $encoding;
        } catch (Exception $ex) {
            $this->error = $ex->getMessage();
            return("Linea: ".__LINE__ . ": " . $ex->getMessage());
        }
    }
    
    public function nuevaHoja($hoja=0, $titulo=null){
        $this->titulo = "Utilice el método setTitulo(\$titulo)";
        $this->subtitulo = "";            
        $this->error = null;              
        $this->html = "";               
        $this->columnas = array();
    }
    
    public function setEncoding($encoding){
         $this->encoding = $encoding;
    }
    
    public function setTitulo($titulo){         
        $this->titulo = ($titulo)? "<div class='titulo'> <h2>{$titulo}</h2></div>": "";            
    }
    
    public function setSubTitulo($subtitulo){
        $this->subtitulo = ($subtitulo)? "<div class='subtitulo'> <h3>{$subtitulo}</h3></div>": "";            
    }
           
    public function setColumna($columna){
        $this->columnas[] = $columna; 
    }
     
    public function setColumnaMerge($columna){
        
    }
    
    public function setEncabezadoMerge(){
        
    }
    
    private function setHTML($html){
        $this->html .= $html;
    }
    
    public function setEncabezados(){
        try{
            $encabezados = $this->columnas;
            $cols = count($encabezados);

            if ($cols == 0) {
                throw new Exception("No se ha configurado los encabezados");
            }
            
            $thead = "<tr>";
            $i = 0;
            foreach ($encabezados as $encabezado) {
                $thead .= "<th width='{$encabezado["ancho"]}' align='{$encabezados[$i++]["alineacion"]}'>{$encabezado["nombre"]}</th>";                
            }
            $thead .= "</tr>";
            
            return($thead);
            
        } catch (Exception $ex) {
            $this->error = $ex->getMessage();
            return("Linea: ".__LINE__ . ": " . $ex->getMessage());   
        }
        
    }
    
    private function setTotales($data, $opciones){
        try{
            $encabezados = $this->columnas;
            $cols = count($encabezados);
            $totales= array();        
            $j = 0;
            $t = 0;
            $tfoot = "<tr>";        
            
            for($i=0; $i< $cols; $i++){                
                if ($encabezados[$i]["totalizar"] == true) {
                    $totales[$i] = array_sum(array_column($data, $i));
                    $t++;
                }
                else{
                    $totales[$i] = null;
                    if($i==0){
                        $totales[0] = "TOTALES:";
                        $j=1;
                        $tfoot .= "<td><b>TOTALES:</b></td>";
                    }
                }
            }

            for ($i = $j; $i < $cols; $i++) {
                if (is_numeric($totales[$i])) {
                    $decimales = (isset($encabezados[$i]["decimales"]))? $encabezados[$i]["decimales"]: 0;
                    $campo = number_format($totales[$i], $decimales);
                    $tfoot .= "<td align='right'>{$campo}</td>";
                } else {
                    $tfoot .= "<td>&nbsp;</td>";
                }
            }
            $tfoot .= "</tr>";            
            if($t == 0){
                $tfoot = "";
            }
            return($tfoot);
        } catch (Exception $ex) {
            $this->error = $ex->getMessage();
            return("Linea: ".__LINE__ . ": " . $ex->getMessage());   
        }        
    }
    
    public function setReporte($datos, $opciones = array()) {
        try {
            $encabezados = $this->columnas;
            
            if (count($encabezados) == 0) {
                throw new Exception("No se ha configurado los encabezados");
            }
            
            $this->data = $datos;
            $thead = $this->setEncabezados();
            $tbody = "";
            $tfoot = $this->setTotales($datos, $opciones);
           
            sleep(1);  //En caso de que se genere mas de una tabla, se pausa para que el id no se repita
            $id = time();
            
            foreach ($datos as $data) {
                $tbody .= "<tr>";
                $col = 0;
                foreach ($data as $celda) {                                        
                    switch(strtolower($encabezados[$col]["tipo"])){
                        case "entero":
                            $campo = is_numeric($celda)? number_format($celda, 0, ",", ".") : $celda;
                        break;
                        case "decimal":
                            $decimales = (isset($encabezados[$col]["decimales"]))? $encabezados[$col]["decimales"] : 2;
                            $campo = is_numeric($celda)? number_format($celda, $decimales, ",", ".") : $celda;
                        break;
                        case "moneda":
                            $campo = is_numeric($celda)? "$".number_format($celda, 2, ",", ".") : $celda; 
                        break;
                        case "fecha":
                            /*$fecha = strtotime($celda);                            
                            $campo = date('d/m/Y', $fecha);*/
                            $campo = $celda;
                        break;
                        default: //texto
                            $campo = (strtolower($this->encode) != "utf8") ? utf8_encode($celda) : $celda;
                    }
                    $tbody .= "<td align='{$encabezados[$col]["alineacion"]}'>{$campo}</td>";
                    $col++;
                }                
                $col = 0;
                $tbody .= "</tr>";
            }
            
            $width = "100%"; 
            $leyenda = "";
            if(isset($opciones['leyenda'])){
                $width = "85%";
                $leyenda = $this->setLeyenda($opciones['leyenda']);
            }
                        
            $style = (isset($opciones['estilo-tabla']))? "style='{$opciones['estilo-tabla']}'" : "";
            $class = (isset($opciones['class-tabla']))? "class='{$opciones['class-tabla']}'" : "";
            
            $width = (isset($opciones['ancho']))? $opciones['ancho'] : "100%";
            $height = (isset($opciones['alto']))? $opciones['alto'] : "600px";
            
            $html = <<<HTML
                <div style="width:{$width}; height:{$height}; overflow:auto;">    
                    {$this->titulo}
                    {$this->subtitulo}
                    <div style="padding: 10px; float: left;">
                        <table id='tabla_reporte_{$id}' {$class} {$style} cellspacing='0'>
                            <thead>
                                {$thead}
                            </thead>
                            <tfoot>
                                {$tfoot}
                            </tfoot>
                            <tbody>
                                {$tbody}
                            </tbody>
                        </table>
                    </div>        
                    {$leyenda}
                </div    
HTML;
            
            //Si en las opciones se seteo para que utilizara el JQuery DataTable
            if(isset($opciones["data_table"]) && $opciones["data_table"] == true){
                $html .= $this->applyDataTable('tabla_reporte_'.$id,$opciones);
            }
            
            $this->setHTML($html);
        } catch (Exception $ex) {
            $this->error = $ex->getMessage();
            echo ("Linea: ".__LINE__ . ": " . $ex->getMessage()."<br/>");
        }
    }
    
    public function setLeyenda($info){
        $leyenda = "<div style='float:left; margin-top: 39px; padding: 10px; width:15%'>";
        $leyenda .= "<table class='leyenda'>";        
        foreach($info as $fila){
            $leyenda .= "<tr>"; 
            foreach($fila as $col){
                $leyenda .= "<td>{$col}</td>"; 
            }        
            $leyenda .= "<tr>"; 
        }
        $leyenda .= "</table>";
        $leyenda .= "</div>";
        return($leyenda);
    }
    
    private function applyDataTable($tabla, $opciones){
        $ordena = (!empty($opciones["ordenar"]))? $opciones["ordenar"]: true;
        $buscar = (!empty($opciones["buscar"]))? $opciones["buscar"]: true;
        $pagina = (!empty($opciones["paginar"]))? $opciones["paginar"]: true;
        $cant_elementos_pag = (!empty($opciones["cant_elementos_pag"]))? $opciones["cant_elementos_pag"]: 10;

        return <<<HTML
           <script type="text/javascript">
               $(document).ready( function () {
                   $('#{$tabla}').DataTable({
                        responsive: true,
                        language: {
                            "url": "../vendor/DataTables/Spanish.json"
                        }
                   });
               });   
           </script>
HTML;

    }

    public function setGrafico($data){
        $titulo = $data["titulo"];
        $ancho = $data["dimensiones"]["ancho"];
        $alto = $data["dimensiones"]["alto"];
        $abscisa = $data["abscisa"];
        $cols = $data["datos"]["columnas"][0];
        $datos = $data["datos"];
        
        if(isset($data["datos"]["filas"]["filas_ini"])){
            $fila_ini = $data["datos"]["filas"]["filas_ini"];
            if(isset($data["datos"]["filas"]["filas_fin"])){
                $fila_fin = $data["datos"]["filas"]["filas_fin"];
            }
        }
        else{
            $fila_ini = 0;
            $fila_fin = count($datos);
        }

        $rows = "";
        for($i = $fila_ini; $i < $fila_fin; $i++){
            $rows .= "['{$datos[$i][$abscisa]}',{$datos[$i][$cols]}],";
        }


        return <<<HTML
        <script type="text/javascript">

            // Load the Visualization API and the corechart package.
            google.charts.load('current', {'packages':['corechart']});

      // Set a callback to run when the Google Visualization API is loaded.
      google.charts.setOnLoadCallback(drawChart);

      // Callback that creates and populates a data table,
      // instantiates the pie chart, passes in the data and
      // draws it.
      function drawChart() {

          // Create the data table.
          var data = new google.visualization.DataTable();
          data.addColumn('string', 'Topping');
          data.addColumn('number', 'Slices');          
          data.addRows([
              $rows
          ]);

          // Set chart options
          var options = {'title':'{$titulo}',
                       'width':{$ancho},
                       'height':{$alto}};

        // Instantiate and draw our chart, passing in some options.
        var chart = new google.visualization.PieChart(document.getElementById('chart_div'));
        chart.draw(data, options);
      }
    </script>
    <!--Div that will hold the pie chart-->
    <div id="chart_div"></div>
HTML;
    }
    
        public function show($tipo=null){            
            return $this->html;            
        }
}
