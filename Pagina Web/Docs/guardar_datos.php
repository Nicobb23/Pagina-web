<?php

// Incluimos la clase PhpSpreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Comprobamos si se han enviado los datos del formulario
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Recogemos los datos del formulario
    $nombre = $_POST["nombre"];
    $edad = $_POST["edad"];
    $correo = $_POST["correo"];
    $comentario = $_POST["comentario"];

    // Creamos una instancia de PhpSpreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Añadimos los encabezados
    $sheet->setCellValue('A1', 'Nombre');
    $sheet->setCellValue('B1', 'Edad');
    $sheet->setCellValue('C1', 'Correo');
    $sheet->setCellValue('D1', 'Comentario');

    // Añadimos los datos
    $row = 2;
    $sheet->setCellValue('A' . $row, $nombre);
    $sheet->setCellValue('B' . $row, $edad);
    $sheet->setCellValue('C' . $row, $correo);
    $sheet->setCellValue('D' . $row, $comentario);

    // Guardamos el archivo Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save('Pagina web/DocS/formulario_respuestas.xlsx');

    // Redireccionamos de vuelta al formulario
    header("Location: Formulario.html");
    exit();
}

?>