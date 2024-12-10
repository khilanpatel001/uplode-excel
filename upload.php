<?php
session_start();
require 'src/Spout/Autoloader/autoload.php';

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;

$host = 'localhost';
$db = 'report';
$user = 'root';
$pass = '';

$conn = new mysqli($host, $user, $pass, $db);

if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excel-file'])) {
    $fileName = $_FILES['excel-file']['name'];
    $fileTmpName = $_FILES['excel-file']['tmp_name'];
    $fileError = $_FILES['excel-file']['error'];

    // Check for upload errors
    if ($fileError !== UPLOAD_ERR_OK) {
        $_SESSION['upload_error'] = "File upload error: " . $fileError; // Store error message in session
        header("Location: index.php"); // Redirect to index.php
        exit();
    }

    // Read the Excel file
    $reader = ReaderEntityFactory::createXLSXReader();
    $reader->open($fileTmpName);

    $rowCount = 0; // Initialize row counter

    foreach ($reader->getSheetIterator() as $sheet) {
        foreach ($sheet->getRowIterator() as $row) {
            $rowCount++; // Increment the row counter
            // Skip the first row (header)
            if ($rowCount === 1) {
                continue;
            }

            $rowData = $row->getCells();

            // Retrieve values from the row
            $userId = isset($rowData[0]) ? $rowData[0]->getValue() : '';
            $userName = isset($rowData[1]) ? $rowData[1]->getValue() : '';
            $department = isset($rowData[2]) ? $rowData[2]->getValue() : '';
            $subDepartment = isset($rowData[3]) ? $rowData[3]->getValue() : '';
            $workshop = isset($rowData[4]) ? $rowData[4]->getValue() : '';
            $role = isset($rowData[5]) ? $rowData[5]->getValue() : '';
            $line = isset($rowData[6]) ? $rowData[6]->getValue() : '';
            $category = isset($rowData[7]) ? $rowData[7]->getValue() : '';
            $workCategory = isset($rowData[8]) ? $rowData[8]->getValue() : '';
            $schShiftId = isset($rowData[9]) ? $rowData[9]->getValue() : '';
            $inTime = isset($rowData[10]) ? $rowData[10]->getValue() : '';
            $outTime = isset($rowData[11]) ? $rowData[11]->getValue() : '';
            $workHrs = isset($rowData[12]) ? $rowData[12]->getValue() : '';
            $actualWorkHrs = isset($rowData[13]) ? $rowData[13]->getValue() : '';
            $overtime = isset($rowData[14]) ? $rowData[14]->getValue() : '';

            // Convert In Time and Out Time to MySQL TIME format (if necessary)
            $inTimeFormatted = null;
            $outTimeFormatted = null;

            // Handle In Time
            if ($inTime instanceof \DateTime) {
                $inTimeFormatted = $inTime->format('H:i:s'); // Convert to MySQL TIME format (HH:MM:SS)
            } elseif (is_string($inTime) || is_numeric($inTime)) {
                $inTimeFormatted = date('H:i:s', strtotime($inTime)); // Convert string to MySQL TIME format
            }

            // Handle Out Time
            if ($outTime instanceof \DateTime) {
                $outTimeFormatted = $outTime->format('H:i:s'); // Convert to MySQL TIME format (HH:MM:SS)
            } elseif (is_string($outTime) || is_numeric($outTime)) {
                $outTimeFormatted = date('H:i:s', strtotime($outTime)); // Convert string to MySQL TIME format
            }

            // Prepare and execute SQL statement (with just time)
            $stmt = $conn->prepare("INSERT INTO attendance_report (User_ID, User_Name, Department, Sub_Department, Workshop, Role, Line, Category, Work_Category, Sch_Shift_ID, In_Time, Out_Time, Work_Hrs, ACTUAL_WORK_HRS, OVERTIME) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

            if ($stmt === false) {
                error_log("Prepare failed: " . $conn->error);
                continue; // Skip to next row if preparation fails
            }

            // Bind parameters with In_Time and Out_Time as TIME
            $stmt->bind_param("sssssssssssssss", $userId, $userName, $department, $subDepartment, $workshop, $role, $line, $category, $workCategory, $schShiftId, $inTimeFormatted, $outTimeFormatted, $workHrs, $actualWorkHrs, $overtime);

            // Execute the statement and check for errors
            if (!$stmt->execute()) {
                error_log("SQL Error: " . $stmt->error); // Log any SQL errors
            }
        }
    }

    $reader->close();
    echo "<script>alert('Excel data uploaded successfully!'); window.location.href='index.php';</script>";
} else {
    echo "Please upload an Excel file.";
}

$conn->close();
?>