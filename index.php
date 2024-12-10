<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Attendance Excel Upload</title>
    <link rel="stylesheet" href="styles.css">
</head>

<body>
    <div class="container">
        <h1>Upload Attendance Excel Sheet</h1>
        <form action="upload.php" method="post" enctype="multipart/form-data">
            <label for="excel-file">Select Excel file:</label>
            <input type="file" name="excel-file" id="excel-file" accept=".xls, .xlsx" required>
            <br>
            <button type="submit">Upload</button>
        </form>
        <br>
        <!-- Updated link to point to sampleexcel.php -->
        <a href="sampleexcel.php">
            <button>Download Sample Excel Sheet</button>
        </a>
    </div>
</body>

</html>