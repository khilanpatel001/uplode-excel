<?php
require 'src/Spout/Autoloader/autoload.php'; // Load Spout's autoloader

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Type;

// Create a new writer for XLSX format
$writer = WriterEntityFactory::createXLSXWriter();
$writer->openToBrowser('sample_work_hours.xlsx'); // Specify the filename for download


$headerRow = WriterEntityFactory::createRowFromArray([
    'User ID',
    'User Name',
    'Department',
    'Sub Department',
    'Workshop',
    'Role',
    'Line',
    'Category',
    'Work Category',
    'Sch Shift ID',
    'In Time',
    'Out Time',
    'Work Hrs',
    'ACTUAL WORK HRS',
    'OVERTIME'
]);


$writer->addRow($headerRow);

$sampleRow1 = WriterEntityFactory::createRowFromArray([
    '1', 
    'John Doe', 
    'Engineering', 
    'Mechanical', 
    'Main Workshop', 
    'Engineer', 
    'Line 1', 
    'Production', 
    'Daily', 
    'S1', 
    '08:00', 
    '05:00', 
    '9', 
    '9', 
    '0' 
]);


$writer->addRow($sampleRow1);


$sampleRow2 = WriterEntityFactory::createRowFromArray([
    '2', 
    'Jane Smith', 
    'Marketing', 
    'Digital', 
    'Creative Workshop', 
    'Manager', 
    'Line 2', 
    'Marketing', 
    'Weekly', 
    'S2', 
    '09:00', 
    '06:00', 
    '9', 
    '8.5', 
    '0.5' 
]);

// Write the second sample row to the file
$writer->addRow($sampleRow2);

// Close the writer
$writer->close();
?>