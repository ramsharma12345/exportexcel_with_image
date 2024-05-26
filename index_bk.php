<?php
require_once("../../config.php");

// Include necessary files
require_once($CFG->libdir.'/excellib.class.php');

// Create a new workbook
$workbook = new MoodleExcelWorkbook('my_data.xlsx');

// Optional: Add a worksheet with a name
$worksheet = $workbook->add_worksheet('My Data');

// Sample data (replace with your actual data retrieval logic)
$data = array(
  array('Name', 'Email', 'Grade','URL'),
  array('John Doe', 'john.doe@example.com', 85, IMAGE('http://localhost/moodle4.4/pluginfile.php/1/core_admin/logocompact/300x300/1716569410/pexels-cottonbro-4065902.jpg')), 
);

// Write data to the worksheet
$row = 0;
foreach ($data as $record) {
  $col = 0;
  foreach ($record as $field) {
    $worksheet->write($row, $col, $field);
    $col++;
  }
  $row++;
}

// Close the workbook and send it to the browser as a download
$workbook->close();
$filename = 'my_data2.xlsx';
$workbook->send($filename);

?>
