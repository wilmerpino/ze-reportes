# ze-reportes
Libreria para generar reportes en Excel y HTML

== Excel ==

Archivos a incluir en el formulario
<pre>
include dirname($_SERVER['DOCUMENT_ROOT']) . '/include/base/PHPExcel-1.8/Classes/PHPExcel.php';
include dirname($_SERVER['DOCUMENT_ROOT']) . "/include/base/ze-reportes/excel-reportes.php";
include dirname($_SERVER['DOCUMENT_ROOT']) . "/include/base/ze-reportes/html-reportes.php";
include dirname($_SERVER['DOCUMENT_ROOT']) . "/include/base/ze-reportes/styles.php";
</pre>



'''Constructor'''
<pre>
public function __construct($encoding = "UTF8", $propiedades = array())

$reporte = new excel_reportes("UTF8", $propiedades);
La codificación por defecto es UTF8 de manera que todos los texto que se envían a la clase se convierte a utf8
 
</pre>

'''Configuración de los parámetros'''

<pre>
$propiedades = array(
     "nombre_archivo" => "ejemplo.xlsx",
     "creado_por" => "WIlmer Pino",
     "modificado_por" => "Wilmer Pino",
     "asunto" => "Prueba clase excel",
     "titulo" => "Título del archivo",
     "descripcion" => "Sin descripción",
     "claves" => "prueba, excel, php",
     "categoria" => "excel"
);

</pre>

'''Configuración de las columnas '''

Para configurar las cabeceras de las columnas se utilizan dos métdodos setColumna($columna) y setColumnaMerge($merge), el primero configura los encabezados de las columnas y el segundo permite agrupar columnas para mostrar encabezados similares.


----

''Método setColumn($columna)''

<pre>
$reporte->setColumna("nombre" => "RUT", "ancho" => "20", "tipo" => "texto", "alineacion" => "izquierda", "totalizar" => false);
$reporte->setColumna("nombre" => "Nombre", "ancho" => "20", tipo" => "texto", "alineacion" => "izquierda", "totalizar" => false);
$reporte->setColumna("nombre" => "Apellido Paterno", "ancho" => "20",  "tipo" => "texto", "alineacion" => "izquierda", "totalizar" => false);
$reporte->setColumna("nombre" => "Fecha Inicio Contrato",  "ancho" => "10",  "tipo" => "fecha", "alineacion" => "izquierda", "totalizar" => false);
$reporte->setColumna("nombre" => "Monto Horas Extras",  "ancho" => "10", "tipo" => "moneda", "alineacion" => "derecha", "totalizar" => false);

</pre>

''nombre ''
Es el nombre que aparecerá en la cabecera de la columna

''ancho''
El ancho en pixeles de la columna

''tipo''
Se definen 5 tipos de dato, se realizará la converisión al tipo indicado, si no requiere convertor el tipo de dato utilice ''texto''
texto
entero
decimal (para la versión HTML puede especificar con el parámetro ''decimales'' la cantidad de decimales que se mostrarán
moneda
fecha

''alineacion''
Establece la alineación en el archivo excel

''totalizar''
Inidica si la columna va a totalizar al final de la tabla


----

''Método setColumnaMerge($columnas)''

<pre>
$encabezados = array(
        array("nombre" => "Diario", "desde" => 18, "hasta" => 19, "alineacion" => "centro"),
        array("nombre" => "Semanal", "desde" => 20, "hasta" => 21, "alineacion" => "centro"),
        array("nombre" => "Mensual", "desde" => 22, "hasta" => 23, "alineacion" => "centro"),
    ),

$reporte->setColumnaMerge($encabezados);
</pre>

''nombre'' 
nombre que aparece en la cabecera agrupada

''desde y hasta''
indican las columnas que harán merge, desde cuál hasta cuál..

''aliniacion''
Se definen los tres tipo, derecha, izquierda y centro


'''Datos'''

Los datos debe ser pasados al método como un arreglo asociativo

<pre>
$data = array(
   array(
         "rut"        => 25000000-1,
         "nombre"     => "Wilmer",
         "apellido"   => "Pino", 
         "sueldo"     => 4000000,
         "fecha_nac"  => "12/08/1974"
   ) 
);
</pre>

'''Títulos y Sub-Títulos'''
Para establecer los títulos y subtítulos se utilizan dos métdos, setTitulo() y setSubTitulo()

<pre>
$reporte->setTitulo("Título del reporte");
$reporte->setSubTitulo("Sub Título del reporte");
</pre>

''' Construir el reporte '''
Se debe invocar el método setReporte($data) para enviar los datos al reporte, éste debe ser invocado luego de los títulos y subtítulos

<pre>
 $reporte->setReporte($data);
</pre>

'''Leyenda'''
Permite escribir una tabla con una leyenda, puede ubicar la posición de la tabla de leyenda con las constantes LEYENDA_DERECHA (Ubica la leyenda a la derecha de la tabla de datos) y LEYENDA_DEBAJO (ubica la leyenda debajo de la tabla de datos)

<pre>
$leyenda => array(
    'encabezados' => array(
         array("nombre" => "Código", "ancho" => 20),
         array("nombre" => "Valor", "ancho" => 20)
     ),
    'datos' => array(
        array("M", "Maculino"), 
        array("F", "Femenino")
    )
);

$reporte->setLeyenda($leyenda, $reporte::LEYENDA_DERECHA);

</pre>


'''Varias hojas'''

Luego de configurar la primera tabla puede utilizar el método nuevaHoja() para configurar una nueva tabla en una hoja distinta, repitiendo los pasos anteriores

<pre>
public function nuevaHoja($hoja=0, $titulo=null);

$reporte->nuevaHoja(1, "hoja2");
</pre>

''hoja''
Indica la hoja a crear, 0 es la primera hoja

''titulo''
Nombre de la nueva hoja

'''Mostrar la tabla'''
Para mostrar la tabla se debe invocar el método show()

<pre>
$reporte->show();
</pre>


----
