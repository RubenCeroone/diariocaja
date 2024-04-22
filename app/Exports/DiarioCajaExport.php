<?php

namespace App\Exports;

use App\Models\BoxDiaries;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use DateTime;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use Carbon\Carbon;

class BoxDiariesExport implements FromCollection, WithStyles
{
    protected $diarioCajas;

    public function __construct()
    {
        $this->diarioCajas = BoxDiaries::all();
        $this->diarioCajas->prepend(new BoxDiaries());
    }

    public function collection(): Collection
    {
        return $this->diarioCajas;
    }

    public function clearRow1(Worksheet $sheet)
    {
        // Borra los datos de la línea 1 desde la columna C hasta la columna V
        for ($column = 'C'; $column <= 'V'; $column++) {
            $sheet->setCellValue($column . '1', '');
        }
    }
    public function clearRow2(Worksheet $sheet)
    {
        // Borra los datos desde la columna A5 hasta la columna A50
        for ($row = 5; $row <= 2000; $row++) {
            $sheet->setCellValue('A' . $row, '');
        }
    }
    public function clearRow3(Worksheet $sheet)
    {
        // Columnas que queremos limpiar desde la A hasta la S, incluyendo desde la D
        $columns = array_merge(['A'], range('D', 'S'));
        // Fila en la que estamos trabajando
        $row = 4;
        // Borra los datos en las celdas específicas
        foreach ($columns as $column) {
            $sheet->setCellValue($column . $row, '');
        }
    }
    public function clearRow4(Worksheet $sheet)
    {
        // Filas y columnas que queremos limpiar
        $startRow = 2;
        $endRow = 2000;
        $startColumn = 'T';
        $endColumn = 'W';
        // Borra los datos en las celdas específicas
        for ($row = $startRow; $row <= $endRow; $row++) {
            for ($column = $startColumn; $column <= $endColumn; $column++) {
                $sheet->setCellValue($column . $row, '');
            }
        }
    }
    public function addBottomBorder(Worksheet $sheet, $startRow, $endRow, $startColumn, $endColumn)
    {
        // Encontrar la última fila donde se insertaron datos
        $lastRow = $endRow;
        while ($lastRow >= $startRow && empty(trim($sheet->getCell($startColumn . $lastRow)->getValue()))) {
            $lastRow--;
        }

        // Establecer el rango de celdas para aplicar el borde
        $range = $startColumn . $lastRow . ':' . $endColumn . $lastRow;

        // Aplicar el borde con el color #318CE7
        $sheet->getStyle($range)->applyFromArray([
            'borders' => [
                'bottom' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '318CE7'],
                ],
            ],
        ]);

        // Agregar "Total:" en negrita en la columna B
        $sheet->getStyle('B' . ($lastRow + 1))->getFont()->setBold(true);
        $sheet->getStyle('B' . ($lastRow + 1))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue('B' . ($lastRow + 1), "Total:");

        // Crear un bucle para iterar sobre las columnas de D a S
        for ($column = 'D'; $column <= 'S'; $column++) {
            // Calcular la suma de la columna actual
            $sumFormula = '=SUM(' . $column . ($startRow) . ':' . $column . ($lastRow) . ')';
            // Colocar la fórmula de suma en la celda debajo de la línea negra
            $sheet->setCellValue($column . ($lastRow + 1), $sumFormula);

            // Cambiar el color de las cifras en la columna O según su valor
            if ($column == 'O') {
                $totalValue = $sheet->getCell($column . ($lastRow + 1))->getCalculatedValue();

                if ($totalValue > 0) {
                    $sheet->getStyle($column . ($lastRow + 1))->getFont()->getColor()->setARGB('008000'); //verde
                } elseif ($totalValue < 0) {
                    $sheet->getStyle($column . ($lastRow + 1))->getFont()->getColor()->setARGB('FF0000'); // rojo
                } else {
                    $sheet->getStyle($column . ($lastRow + 1))->getFont()->getColor()->setARGB('000000'); //negro
                }
            }
            $sheet->getStyle($column . ($lastRow + 1))->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
        }
    }

    public function colorearFinSemana($sheet, $startRow, $endRow) {
        for ($row = $startRow; $row <= $endRow; $row++) {
            // Obtener el día de la semana desde la celda B$row
            $dia = $sheet->getCell('B' . $row)->getValue();
    
            // Obtener la fecha desde la celda C$row
            $fecha = $sheet->getCell('C' . $row)->getValue();
    
            // Convertir la fecha a un objeto DateTime
            $fecha_obj = new DateTime($fecha);
    
            // Obtener el nombre del día de la semana (en inglés)
            $nombre_dia = $fecha_obj->format('l');
    
            // Mapear el nombre del día de la semana a su equivalente en español
            $dias_semana = [
                'Monday' => 'Lunes',
                'Tuesday' => 'Martes',
                'Wednesday' => 'Miércoles',
                'Thursday' => 'Jueves',
                'Friday' => 'Viernes',
                'Saturday' => 'Sábado',
                'Sunday' => 'Domingo'
            ];
    
            // Obtener el nombre del día de la semana en español
            $nombre_dia_espanol = $dias_semana[$nombre_dia];
    
            // Poner el nombre del día de la semana en la celda B$row
            $sheet->setCellValue('B' . $row, $nombre_dia_espanol);
    
            // Si el día es sábado o domingo, colorear fondo en gris claro desde B hasta S
            if ($nombre_dia === 'Saturday' || $nombre_dia === 'Sunday') {
                $sheet->getStyle('B' . $row . ':S' . $row)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => [
                            'rgb' => 'D3D3D3', // Gris claro
                        ],
                    ],
                ]);
            }
        }
    }

    public function styles(Worksheet $sheet)
    {
        // Combina las celdas A1 y B1
        $sheet->mergeCells('A1:B1');
        // Establece el texto "Diario de Caja" en las celdas combinadas
        $sheet->setCellValue('A1', 'Diario de Caja');
        // Poner tamaño para el texto
        $sheet->getStyle('A1')->getFont()->setSize(22);
        // Aplica estilos al texto "Diario de Caja"
        $sheet->getStyle('A1')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Cambiar el color de fondo a azul claro para las celdas A2, A3, y N5
        $lightBlueStyle = [
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => '318CE7'], // Azul claro, por ejemplo 'Sky Blue'
            ],
        ];
        $sheet->getStyle('A2')->applyFromArray($lightBlueStyle);
        $sheet->getStyle('A3')->applyFromArray($lightBlueStyle);
        // Aplicar el mismo estilo a cualquier otra celda que necesite el fondo azul claro
        // ...
        // Cambiar el estilo de fondo para "Empresas de Seguridad" a azul claro
        $sheet->getStyle('Q5')->applyFromArray($lightBlueStyle);
        $sheet->getStyle('C1:R1')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'FFFFFF', // Color blanco
                ],
            ],
        ]);
        // Llamar a la función clearRow1
        $this->clearRow1($sheet);
        // Llamar a la función clearRow2
        $this->clearRow2($sheet);
        // Llamar a la función clearRow3
        $this->clearRow3($sheet);
        // Llamar a la función clearRow4
        $this->clearRow4($sheet);
        // Combina las celdas A2 y R2
        $sheet->mergeCells('A2:S2');
        // Establece el texto "Centro:" en las celdas combinadas
        foreach ($this->collection() as $diarioCaja) {
            // Establece el texto "Centro:" en las celdas combinadas
            $sheet->setCellValue('A2', 'Centro: '. $diarioCaja->center);
        }
        // Aplica estilos al texto "Centro:"
        $sheet->getStyle('A2')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'], // Color blanco
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => '318CE7'], // Color azul oscuro
            ],
        ]);
        // Combina las celdas A3 y R3
        $sheet->mergeCells('A3:S3');
        // Establece el texto "Centro:" en las celdas combinadas
        $sheet->setCellValue('A3', 'Periodo:');
        // Aplica estilos al texto "Centro:"
        $sheet->getStyle('A3')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => 'FFFFFF'], // Color blanco
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => '318CE7'], // Color azul oscuro
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
                'startColor' => ['argb' => '318CE7'], // Color azul oscuro
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
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Dia" en las celdas
        $sheet->setCellValue('B5', 'Dia de la Semana');
        // Aplica estilos al texto "Dia"
        $sheet->getStyle('B5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Fecha" en las celdas
        $sheet->setCellValue('C5', 'Fecha');
        // Aplica estilos al texto "Fecha"
        $sheet->getStyle('C5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Venta REVO (IVA incluido)" en las celdas
        $sheet->setCellValue('D5', 'Venta REVO (IVA incluido)');
        // Aplica estilos al texto "Venta REVO (IVA incluido)"
        $sheet->getStyle('D5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
        ]);
        // Establece el texto "Caja Fuerte Inicio" en las celdas
        $sheet->setCellValue('E5', 'Caja Fuerte Inicio');
        // Aplica estilos al texto "Caja Fuerte Inicio"
        $sheet->getStyle('E5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
        ]);
        // Establece el texto "Efectivo Diario" en las celdas
        $sheet->setCellValue('F5', 'Efectivo Diario');
        // Aplica estilos al texto "Efectivo Diario"
        $sheet->getStyle('F5')->applyFromArray([
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
        $sheet->setCellValue('G5', 'Tarjetas');
        // Aplica estilos al texto "Tarjetas"
        $sheet->getStyle('G5')->applyFromArray([
            'font' => [
                'color' => ['argb' => 'FF0000'], // Color rojo
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "CoverManager" en las celdas
        $sheet->setCellValue('H5', 'CoverManager');
        // Aplica estilos al texto "CoverManager"
        $sheet->getStyle('H5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Transferencia" en las celdas
        $sheet->setCellValue('I5', 'Transferencia');
        // Aplica estilos al texto "Transferencia"
        $sheet->getStyle('I5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Hotel" en las celdas
        $sheet->setCellValue('J5', 'Hotel');
        // Aplica estilos al texto "Hotel"
        $sheet->getStyle('J5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Guestpro" en las celdas
        $sheet->setCellValue('K5', 'Guestpro');
        // Aplica estilos al texto "Guestpro"
        $sheet->getStyle('K5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Pendiente/Cobro" en las celdas
        $sheet->setCellValue('L5', 'Pendiente/Cobro');
        // Aplica estilos al texto "Pendiente/Cobro"
        $sheet->getStyle('L5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true
            ],
        ]);
        // Establece el texto "Propinas" en las celdas
        $sheet->setCellValue('M5', 'Propinas');
        // Aplica estilos al texto "Propinas"
        $sheet->getStyle('M5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Otras Formas Pago" en las celdas
        $sheet->setCellValue('N5', 'Otras Formas Pago');
        // Aplica estilos al texto "Otras Formas Pago"
        $sheet->getStyle('N5')->applyFromArray([
            'font' => [
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'wrapText' => true, // Activa el ajuste de texto
            ],
        ]);
        // Establece el texto "Exceso/Quebranto" en las celdas
        $sheet->setCellValue('O5', 'Exceso/Quebranto');
        // Aplica estilos al texto "Exceso/Quebranto"
        $sheet->getStyle('O5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Retiros" en las celdas
        $sheet->setCellValue('P5', 'Retiros');
        // Aplica estilos al texto "Retiros"
        $sheet->getStyle('P5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Bancos" en las celdas
        $sheet->setCellValue('Q5', 'Bancos');
        // Aplica estilos al texto "Bancos"
        $sheet->getStyle('Q5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        // Establece el texto "Empresas de Seguridad" en las celdas
        $sheet->setCellValue('R5', 'Empresas de Seguridad');
        // Aplica estilos al texto "Empresas de Seguridad"
        $sheet->getStyle('R5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
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
        $sheet->setCellValue('S5', 'Caja Fuerte Final');
        // Aplica estilos al texto "Caja Fuerte Final"
        $sheet->getStyle('S5')->applyFromArray([
            'font' => [
                'bold' => true,
                'color' => ['argb' => '318CE7'], // Color azul
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
        $startRow = 6; // Fila donde empiezan los datos
        $endRow = 2000; // Última fila donde se insertarán datos
        $row = $startRow;
        foreach ($this->collection() as $diarioCaja) {
            if ($row > $endRow) {
                break; // Salir del bucle si alcanzamos la última fila
            }
            $sheet->setCellValue('B' . $row, $diarioCaja->dia);
            $sheet->setCellValue('C' . $row, $diarioCaja->fecha);
            $sheet->setCellValue('D' . $row, $diarioCaja->venta_revo_iva_incluido);
            $sheet->getStyle('D' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('E' . $row, $diarioCaja->caja_fuerte_inicio);
            $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('F' . $row, $diarioCaja->efectivo_diario);
            $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('G' . $row, $diarioCaja->tarjetas);
            $sheet->getStyle('G' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('H' . $row, $diarioCaja->cover_manager);
            $sheet->getStyle('H' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('I' . $row, $diarioCaja->transferencia);
            $sheet->getStyle('I' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('J' . $row, $diarioCaja->hotel);
            $sheet->getStyle('J' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('K' . $row, $diarioCaja->guestpro);
            $sheet->getStyle('K' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('L' . $row, $diarioCaja->pendientes);
            $sheet->getStyle('L' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('M' . $row, $diarioCaja->propinas);
            $sheet->getStyle('M' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('N' . $row, $diarioCaja->otras_formas_pago);
            $sheet->getStyle('N' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('O' . $row, $diarioCaja->exceso_quebranto);
            $sheet->getStyle('O' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('P' . $row, $diarioCaja->retiros);
            $sheet->getStyle('P' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('Q' . $row, $diarioCaja->bancos);
            $sheet->getStyle('Q' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('R' . $row, $diarioCaja->empresas_seguridad);
            $sheet->getStyle('R' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $sheet->setCellValue('S' . $row, $diarioCaja->caja_fuerte_final);
            $sheet->getStyle('S' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
            $row++;
        }
        // Alinear todos los datos a la derecha desde C7 hasta O50
        for ($row = $startRow; $row <= $endRow; $row++) {
            for ($col = 'D'; $col <= 'P'; $col++) {
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
        $fechaCarbon = Carbon::createFromFormat('Y-m-d', $fechaCelda);

        // Obtener el día de la semana como número (0 = domingo, 6 = sábado)
        $diaSemana = $fechaCarbon->dayOfWeek;

        // Si es sábado (6) o domingo (0), no hacer nada
        if ($diaSemana == 6 || $diaSemana == 0) {
            // No aplicar ningún estilo de fondo
        } else {
            // Aplicar otro estilo
        }
        for ($row = $startRow; $row <= $endRow; $row++) {
            $sheet->getStyle('E' . $row)->applyFromArray([
                'font' => [
                    'bold' => true, // Hacer que el texto sea negrita
                ],
            ]);
        }
        for ($row = $startRow; $row <= $endRow; $row++) {
            $valorCelda = $sheet->getCell('O' . $row)->getValue();
            $estilo = [];
            if ($valorCelda > 0) {
                // Si el valor es positivo, establecer fuente azul y negrita
                $estilo = [
                    'font' => [
                        'bold' => true,
                        'color' => ['rgb' => '008000'], // Color verde intermedio
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
            $sheet->getStyle('O' . $row)->applyFromArray($estilo);
        }
        // Ajustar altura de la fila 1
        $sheet->getRowDimension(1)->setRowHeight(30.75);
        // Ajustar altura de la fila 5
        $sheet->getRowDimension(5)->setRowHeight(43.5);
        // Ajustar ancho de las columnas
        $sheet->getColumnDimension('A')->setWidth(8.43);
        $sheet->getColumnDimension('B')->setWidth(18.86);
        $sheet->getColumnDimension('C')->setWidth(18.86);
        $sheet->getColumnDimension('D')->setWidth(15.86);
        $sheet->getColumnDimension('E')->setWidth(12.86);
        $sheet->getColumnDimension('F')->setWidth(13.14);
        $sheet->getColumnDimension('G')->setWidth(15.29);
        $sheet->getColumnDimension('H')->setWidth(13.57);
        $sheet->getColumnDimension('I')->setWidth(13.14);
        $sheet->getColumnDimension('J')->setWidth(13.14);
        $sheet->getColumnDimension('K')->setWidth(13.14);
        $sheet->getColumnDimension('L')->setWidth(13.14);
        $sheet->getColumnDimension('M')->setWidth(11.29);
        $sheet->getColumnDimension('N')->setWidth(8.43);
        $sheet->getColumnDimension('O')->setWidth(16.71);
        $sheet->getColumnDimension('P')->setWidth(11.29);
        $sheet->getColumnDimension('Q')->setWidth(10.29);
        $sheet->getColumnDimension('R')->setWidth(16);
        $sheet->getColumnDimension('S')->setWidth(18.29);
        // Establecer el estilo de borde para eliminar cualquier borde entre las columnas B y C
        $noBorderStyle = [
            'borders' => [
                'right' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE,
                ],
            ],
        ];
        // Aplicar el estilo de borde sin borde a la columna B hasta la fila 10
        $sheet->getStyle('B1:B10')->applyFromArray($noBorderStyle);
        // Establecer el estilo de fuente en negrita
        $fontStyle = [
            'font' => [
                'bold' => true,
            ],
        ];
        // Establecer el estilo de borde para la parte superior e inferior
        $borderStyle = [
            'borders' => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '87CEEB'], // Color azul
                ],
                'bottom' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '87CEEB'], // Color azul
                ],
            ],
        ];
        // Combinar el estilo de fuente y el estilo de borde
        $mergedStyle = array_merge($fontStyle, $borderStyle);
        // Aplicar el estilo combinado a las celdas desde B5 hasta R5
        $sheet->getStyle('B5:S5')->applyFromArray($mergedStyle);

        $this->addBottomBorder($sheet, $startRow, $endRow, 'B', 'S');

        $this->colorearFinSemana($sheet, $startRow, $endRow);
    }
}