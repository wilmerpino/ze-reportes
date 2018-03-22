<?php

$_style = array(
    'bordes' => array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('rgb' => '000000')
            )
        )
    )
    ,
    'titulo' => array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        ),
        'font' => array(
            'bold' => true,
            'name' => 'Verdana',
            'color' => array('rgb' => '000000'),
            'size' => 14
        )
    ),
    'subtitulo' => array(
            'alignment' => array(
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
            ),
            'font' => array(
                'bold' => true,
                'name' => 'Verdana',
                'color' => array('rgb' => '000000'),
                'size' => 12
            ),
    ),
    'encabezados' =>  array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => '62727b'),
            ),
            'inside' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => '62727b'),
            ),
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'wrap' => true
        ),
        'font' => array(
            'bold' => true,
            'name' => 'Verdana',
            'color' => array('rgb' => 'FFFFFF'),
            'size' => 10
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'rotation' => 90,
            'color' => array(
                'argb' => '0277BD'
            ),
            'startColor' => array(
                'argb' => 'FFFFFF',
            ),
            'endColor' => array(
                'argb' => 'FFFFFFFF',
            )
        )
    ),
    'encabezados_agrupados' =>  array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => '62727b'),
            ),
            'inside' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => 'FFFFFF'),
            ),
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'wrap' => true
        ),
        'font' => array(
            'bold' => true,
            'name' => 'Verdana',
            'color' => array('rgb' => 'FFFFFF'),
            'size' => 10
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'rotation' => 90,
            'color' => array(
                'argb' => '004C8C'
            ),
            'startColor' => array(
                'argb' => 'FFFFFF',
            ),
            'endColor' => array(
                'argb' => 'FFFFFFFF',
            )
        )
     ),   
    'encabezados_leyenda' =>  array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => '62727b'),
            ),
            'inside' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => 'FFFFFF'),
            ),
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'wrap' => true
        ),
        'font' => array(
            'bold' => true,
            'name' => 'Verdana',
            'color' => array('rgb' => 'FFFFFF'),
            'size' => 10
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'rotation' => 90,
            'color' => array(
                'argb' => '004C8C'
            ),
            'startColor' => array(
                'argb' => 'FFFFFF',
            ),
            'endColor' => array(
                'argb' => 'FFFFFFFF',
            )
        )
     ),   
    'cuerpo' =>  array(
        'font' => array(
            'bold' => false,
            'name' => 'Verdana',
            'color' => array('rgb' => '000000'),
            'size' => 10
        ),
        'alignment' => array(
            'wrap' => true
        ),
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => '62727b'),
            ),
            'inside' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => '62727b'),
            ),
        ),
    ),
    'fila_impar' =>  array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'rotation' => 90,
            'color' => array(
                'argb' => 'eeeeee'
            ),
            'startColor' => array(
                'argb' => 'FFFFFF',
            ),
            'endColor' => array(
                'argb' => 'FFFFFFFF',
            )
        )
    ),
    'fila_par' =>  array(
        'fill' => array(
            'fillType' => PHPExcel_Style_Fill::FILL_SOLID,
            'rotation' => 90,
            'color' => array(
                'argb' => 'FFFFFF'
            ),
            'startColor' => array(
                'argb' => 'FFFFFF',
            ),
            'endColor' => array(
                'argb' => 'FFFFFFFF',
            ),
        )
    ),
    'titulo_tabla' =>  array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        ),
        'font' => array(
            'bold' => true,
            'name' => 'Verdana',
            'color' => array('rgb' => '000000'),
            'size' => 13
        )
    ),
    'subtitulo_tabla' =>  array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        ),
        'font' => array(
            'bold' => true,
            'name' => 'Verdana',
            'color' => array('rgb' => '000000'),
            'size' => 12
        ),
    )
);


