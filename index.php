<?php
require_once("../../config.php");
require 'vendor/autoload.php';

// MoodleExcelWorkbook

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

// Sample data including image paths
$data = array(
    array('Name', 'Email', 'Grade', 'Image'),
    array('John Doe', 'john.doe@example.com', 85, 'image/ram1.png'),
    array('Jane Smith', 'jane.smith@example.com', 92, 'image/logo.jpg'),
);

// Create a new Spreadsheet object
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Convert column index to letter
function getColumnLetter($index) {
    return chr(65 + $index); // 65 is the ASCII code for 'A'
}

// Add the data to the spreadsheet
foreach ($data as $rowIndex => $row) {
    foreach ($row as $colIndex => $value) {
        if($colIndex == 3 && $value!='Image'){
          continue;
        }
        $cell = getColumnLetter($colIndex) . ($rowIndex + 1);
        $sheet->setCellValue($cell, $value);
    }
}

// Function to add an image to a specified cell
function addImage($worksheet, $path, $coordinates) {
    $drawing = new Drawing();
    $drawing->setName('Image');
    $drawing->setDescription('Image');
    $drawing->setPath($path);
    $drawing->setHeight(50); // Adjust the height as needed
    $drawing->setCoordinates($coordinates);
    $drawing->setWorksheet($worksheet);
}

// Loop through the data to add images
foreach ($data as $rowIndex => $row) {
    if ($rowIndex == 0) continue; // Skip the header row
    $imagePath = $row[3]; // Image path is in the 4th column
    $cellCoordinates = 'D' . ($rowIndex + 1); // Image column is 'D'
    addImage($sheet, $imagePath, $cellCoordinates);
}

// Set headers to prompt for download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="data.xlsx"');
header('Cache-Control: max-age=0');

// Save the spreadsheet to PHP output
$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
 