<?php
 
namespace App\Exports;
 
use App\Models\DiarioCaja;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use Carbon\Carbon;
 
class DiarioCajaExport implements FromCollection, WithStyles
{
    protected $diarioCajas;
 
    public function __construct()
    {
        $this->diarioCajas = DiarioCaja::all();
        $this->diarioCajas->prepend(new DiarioCaja());
    }
 
    public function collection(): Collection
    {
        return $this->diarioCajas;
    }
 
    public function clearRow1(Worksheet $sheet) {
        // Borra los datos de la línea 1 desde la columna C hasta la columna O
        for ($column = 'C'; $column <= 'O'; $column++) {
            $sheet->setCellValue($column . '1', '');
        }
    }
 
    public function styles(Worksheet $sheet)
    {
        // Combina las celdas A1 y B1
        $sheet->mergeCells('A1:B1');
 
        // Establece el texto "Diario de Caja" en las celdas combinadas
        $sheet->setCellValue('A1', 'Diario de Caja');
 
        // Poner tamaño para el texto
        $sheet->getStyle('A1')->getFont()->setSize(18);
 
        // Aplica estilos al texto "Diario de Caja"
        $sheet->getStyle('A1')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
 
        ]);
 
        $sheet->getStyle('C1:O1')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'FFFFFF', // Color blanco
                ],
            ],
        ]);
 
        // Llamar a la función clearRow1
        $this->clearRow1($sheet);
 
        // Combina las celdas A2 y O2
        $sheet->mergeCells('A2:O2');
 
        // Establece el texto "Centro:" en las celdas combinadas
        $sheet->setCellValue('A2', 'Centro:');
 
        // Aplica estilos al texto "Centro:"
        $sheet->getStyle('A2')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'], // Color blanco
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['argb' => '00008B'], // Color azul oscuro
            ],
        ]);
 
        // Combina las celdas A3 y O3
        $sheet->mergeCells('A3:O3');
 
        // Establece el texto "Centro:" en las celdas combinadas
        $sheet->setCellValue('A3', 'Periodo:');
 
        // Aplica estilos al texto "Centro:"
        $sheet->getStyle('A3')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'], // Color blanco
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['argb' => '00008B'], // Color azul oscuro
            ],
        ]);
 
        // Combina las celdas A4 y B4
        $sheet->mergeCells('A4:B4');
 
        // Establece el texto "© MAPAL Software, S.L. Todos los derechos reservados" en las celdas combinadas
        $sheet->setCellValue('A4', '© MAPAL Software, S.L. Todos los derechos reservados');
 
        // Aplica estilos al texto
        $sheet->getStyle('A4')->applyFromArray([
            'font' => [
                'size' => 6,
                'color' => ['argb' => '808080'], // Color gris
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "REVO" en la celda C4
        $sheet->setCellValue('C4', 'REVO');
 
        // Aplica estilos al texto "REVO"
        $sheet->getStyle('C4')->applyFromArray([
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['argb' => 'FFFF00'], // Color amarillo
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Calculation" en la celda E4
        $sheet->setCellValue('E4', 'Calculation');
 
        // Aplica estilos al texto "Calculation"
        $sheet->getStyle('E4')->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "SOLO TARJETAS" en la celda F4
        $sheet->setCellValue('F4', 'SOLO TARJETAS');
 
        // Aplica estilos al texto "SOLO TARJETAS"
        $sheet->getStyle('F4')->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "incluir propinas visa" en la celda I4
        $sheet->setCellValue('I4', 'incluir propinas visa');
 
        // Aplica estilos al texto "incluir propinas visa"
        $sheet->getStyle('I4')->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Este es el acumulado" en la celda O4
        $sheet->setCellValue('O4', 'Este es el acumulado');
 
        // Aplica estilos al texto "Este es el acumulado"
        $sheet->getStyle('O4')->applyFromArray([
            'font' => [
                'color' => ['argb' => 'FF0000'], // Color rojo
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Fecha" en las celdas
        $sheet->setCellValue('B5', 'Fecha');
 
        // Aplica estilos al texto "Fecha"
        $sheet->getStyle('B5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Venta REVO (IVA incluido)" en las celdas
        $sheet->setCellValue('C5', 'Venta REVO (IVA incluido)');
 
        // Aplica estilos al texto "Venta REVO (IVA incluido)"
        $sheet->getStyle('C5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
        ]);
 
        // Establece el texto "Caja Fuerte Inicio" en las celdas
        $sheet->setCellValue('D5', 'Caja Fuerte Inicio');
 
        // Aplica estilos al texto "Caja Fuerte Inicio"
        $sheet->getStyle('D5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
        ]);
 
        // Establece el texto "Efectivo Diario" en las celdas
        $sheet->setCellValue('E5', 'Efectivo Diario');
 
        // Aplica estilos al texto "Efectivo Diario"
        $sheet->getStyle('E5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FF0000'], // Color rojo
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['argb' => 'FFFF00'], // Color amarillo
            ],
        ]);
 
        // Establece el texto "Tarjetas" en las celdas
        $sheet->setCellValue('F5', 'Tarjetas');
 
        // Aplica estilos al texto "Tarjetas"
        $sheet->getStyle('F5')->applyFromArray([
            'font' => [
                'color' => ['argb' => 'FF0000'], // Color rojo
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "CoverManager" en las celdas
        $sheet->setCellValue('G5', 'CoverManager');
 
        // Aplica estilos al texto "CoverManager"
        $sheet->getStyle('G5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Transferencia" en las celdas
        $sheet->setCellValue('H5', 'Transferencia');
 
        // Aplica estilos al texto "Transferencia"
        $sheet->getStyle('H5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Propinas" en las celdas
        $sheet->setCellValue('I5', 'Propinas');
 
        // Aplica estilos al texto "Propinas"
        $sheet->getStyle('I5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Otras Formas Pago" en las celdas
        $sheet->setCellValue('J5', 'Otras Formas Pago');
 
        // Aplica estilos al texto "Otras Formas Pago"
        $sheet->getStyle('J5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
        ]);
 
        // Establece el texto "Exceso/Quebranto" en las celdas
        $sheet->setCellValue('K5', 'Exceso/Quebranto');
 
        // Aplica estilos al texto "Exceso/Quebranto"
        $sheet->getStyle('K5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Retiros" en las celdas
        $sheet->setCellValue('L5', 'Retiros');
 
        // Aplica estilos al texto "Retiros"
        $sheet->getStyle('L5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Bancos" en las celdas
        $sheet->setCellValue('M5', 'Bancos');
 
        // Aplica estilos al texto "Bancos"
        $sheet->getStyle('M5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
 
        // Establece el texto "Empresas de Seguridad" en las celdas
        $sheet->setCellValue('N5', 'Empresas de Seguridad');
 
        // Aplica estilos al texto "Empresas de Seguridad"
        $sheet->getStyle('N5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['argb' => 'FFFF00'], // Color amarillo
            ],
        ]);
 
        // Establece el texto "Caja Fuerte Final" en las celdas
        $sheet->setCellValue('O5', 'Caja Fuerte Final');
 
        // Aplica estilos al texto "Caja Fuerte Final"
        $sheet->getStyle('O5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '0000FF'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['argb' => 'FFFF00'], // Color amarillo
            ],
        ]);
 
        // Insertar datos de DiarioCaja
        $startRow = 7; // Fila donde empiezan los datos
        $endRow = 50; // Última fila donde se insertarán datos
 
        $row = $startRow;
        foreach ($this->collection() as $diarioCaja) {
            if ($row > $endRow) {
                break; // Salir del bucle si alcanzamos la última fila
            }
 
            $sheet->setCellValue('B' . $row, $diarioCaja->fecha);
            $sheet->setCellValue('C' . $row, $diarioCaja->venta_revo_iva_incluido);
            $sheet->getStyle('C' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('D' . $row, $diarioCaja->caja_fuerte_inicio);
            $sheet->getStyle('D' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('E' . $row, $diarioCaja->efectivo_diario);
            $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('F' . $row, $diarioCaja->tarjetas);
            $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('G' . $row, $diarioCaja->cover_manager);
            $sheet->setCellValue('H' . $row, $diarioCaja->transferencia);
            $sheet->setCellValue('I' . $row, $diarioCaja->propinas);
            $sheet->setCellValue('J' . $row, $diarioCaja->otras_formas_pago);
            $sheet->setCellValue('K' . $row, $diarioCaja->exceso_quebranto);
            $sheet->getStyle('K' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('L' . $row, $diarioCaja->retiros);
            $sheet->getStyle('L' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('M' . $row, $diarioCaja->bancos);
            $sheet->setCellValue('N' . $row, $diarioCaja->empresas_seguridad);
            $sheet->getStyle('N' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $sheet->setCellValue('O' . $row, $diarioCaja->caja_fuerte_final);
            $sheet->getStyle('O' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
            $row++;
        }
 
        // Alinear todos los datos a la derecha desde C7 hasta O50
        for ($row = $startRow; $row <= $endRow; $row++) {
            for ($col = 'C'; $col <= 'O'; $col++) {
                $sheet->getStyle($col . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
            }
        }
 
        for ($row = $startRow; $row <= $endRow; $row++) {
            // Obtener el día de la semana de la fecha actual
            /* $fechaCelda = $sheet->getCell('B' . $row)->getValue();
            $fecha = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($fechaCelda);
            $diaSemana = $fecha->format('N');*/
            $fechaCelda = '2022-10-22';
            $fechaCarbon = Carbon::createFromFormat('Y-m-d', $fechaCelda); // Obtener el día de la semana como número (1 = lunes, 7 = domingo)
            $diaSemana = $fechaCarbon->dayOfWeek; //Se refiere al dia de la semana
 
            // Si es sábado o domingo, aplicar fondo blanco a la celda de "Efectivo Diario"
            if ($diaSemana == 6 || $diaSemana == 7) {
                $sheet->getStyle('E' . $row)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => 'FFFFFF', // Color blanco
                        ],
                    ],
                ]);
            }
        }
 
        // Obtener el día de la semana de la fecha actual
        /* $fechaCelda = $sheet->getCell('B' . $row)->getValue();
        $fecha = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($fechaCelda);
        $diaSemana = $fecha->format('N'); */
        $fechaCelda = '2022-10-22';
        $fechaCarbon = Carbon::createFromFormat('Y-m-d', $fechaCelda); // Obtener el día de la semana como número (1 = lunes, 7 = domingo)
        $diaSemana = $fechaCarbon->dayOfWeek; //Se refiere al dia de la semana
 
        // Si es sábado o domingo, aplicar estilo de fondo azul claro a toda la fila
        if ($diaSemana == 6 || $diaSemana == 7) {
            $sheet->getStyle('B' . $row . ':O' . $row)->applyFromArray([
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => [
                        'rgb' => 'E6F3FF', // Color azul claro
                    ],
                ],
        ]);
    }
 
        for ($row = $startRow; $row <= $endRow; $row++) {
            $sheet->getStyle('D' . $row)->applyFromArray([
                'font' => [
                    'bold' => true, // Hacer que el texto sea negrita
                ],
            ]);
        }
 
        for ($row = $startRow; $row <= $endRow; $row++) {
            $valorCelda = $sheet->getCell('K' . $row)->getValue();
            $estilo = [];
 
            if ($valorCelda > 0) {
                // Si el valor es positivo, establecer fuente azul y negrita
                $estilo = [
                    'font' => [
                        'bold' => true,
                        'color' => ['rgb' => '0000FF'], // Color azul
                    ],
                ];
            } elseif ($valorCelda < 0) {
                // Si el valor es negativo, establecer fuente roja y negrita
                $estilo = [
                    'font' => [
                        'bold' => true,
                        'color' => ['rgb' => 'FF0000'], // Color rojo
                    ],
                ];
            } else {
                // Si el valor es cero, establecer fuente negra y negrita
                $estilo = [
                    'font' => [
                        'bold' => true,
                        'color' => ['rgb' => '000000'], // Color negro
                    ],
                ];
            }
 
            // Aplicar estilos a la celda correspondiente en "Exceso/Quebranto"
            $sheet->getStyle('K' . $row)->applyFromArray($estilo);
        }
 
        // Ajustar altura de la fila 1
        $sheet->getRowDimension(1)->setRowHeight(30.75);
 
        // Ajustar altura de la fila 5
        $sheet->getRowDimension(5)->setRowHeight(43.5);
 
        // Ajustar ancho de las columnas
        $sheet->getColumnDimension('A')->setWidth(8.43);
        $sheet->getColumnDimension('B')->setWidth(18.86);
        $sheet->getColumnDimension('C')->setWidth(15.86);
        $sheet->getColumnDimension('D')->setWidth(12.86);
        $sheet->getColumnDimension('E')->setWidth(12.43);
        $sheet->getColumnDimension('F')->setWidth(15.29);
        $sheet->getColumnDimension('G')->setWidth(13.57);
        $sheet->getColumnDimension('H')->setWidth(13.14);
        $sheet->getColumnDimension('I')->setWidth(16.71);
        $sheet->getColumnDimension('J')->setWidth(8.43);
        $sheet->getColumnDimension('K')->setWidth(16.71);
        $sheet->getColumnDimension('L')->setWidth(11.29);
        $sheet->getColumnDimension('M')->setWidth(10.29);
        $sheet->getColumnDimension('N')->setWidth(16);
        $sheet->getColumnDimension('O')->setWidth(18.29);
 
        // Obtener el estilo de borde
        $borderStyle = array(
            'borders' => array(
        '       right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000'), // Color negro
                ),
            ),
        );
 
        // Aplicar el estilo de borde a la columna B hasta la fila 10
        $sheet->getStyle('B1:B10')->applyFromArray($borderStyle);
 
        // Establecer el estilo de fuente en negrita
        $fontStyle = array(
            'font' => array(
                'bold' => true,
            ),
        );
 
        // Establecer el estilo de borde para la parte superior e inferior
        $borderStyle = array(
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '0000FF'), // Color azul
                ),
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '0000FF'), // Color azul
                ),
            ),
        );
 
        // Combinar el estilo de fuente y el estilo de borde
        $mergedStyle = array_merge($fontStyle, $borderStyle);
 
        // Aplicar el estilo combinado a las celdas desde B5 hasta O5
        $sheet->getStyle('B5:O5')->applyFromArray($mergedStyle);
 
    }
}