<?php

require_once('vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;

// Creates New Spreadsheet
$spreadsheet = new Spreadsheet();

// Remove the default sheet
$spreadsheet->removeSheetByIndex(0);

// Create the masterData sheet and populate it with dropdown values
$masterDataSheet = $spreadsheet->createSheet();
$masterDataSheet->setTitle('masterData');

$dropdownValues = [
    ['Kelas', '9A', '9B'],
    ['Kurikulum', 'K13', 'K21'],
    ['Mata Pelajaran', 'Matematika', 'Bahasa Indonesia', 'Bahasa Inggris', 'IPA', 'IPS'],
    ['Semester', 'Ganjil', 'Genap'],
];

$row = 1;
foreach ($dropdownValues as $rowData) {
    $col = 'A';
    foreach ($rowData as $value) {
        $masterDataSheet->setCellValue($col . $row, $value);
        $col++;
    }
    $row++;
}

// Create the importData sheet
$importDataSheet = $spreadsheet->createSheet();
$importDataSheet->setTitle('importData');

$academicYearCode = '2025/2026';

// Set the values for the cells in the importData sheet
$data = [
    ['Tahun Ajar', $academicYearCode],
    [],
    ['Kelas', 'Kurikulum', 'Mata Pelajaran', 'Semester', 'ID Silabus', 'Bab', 'Pokok Pembahasan'],
];

// Set the data in the importData sheet
$row = 1;
foreach ($data as $rowData) {
    $col = 'A';
    foreach ($rowData as $value) {
        $importDataSheet->setCellValue($col . $row, $value);
        $col++;
    }
    $row++;
}

// Generate the columns array dynamically based on the masterData sheet
$columns = [];
$highestColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($masterDataSheet->getHighestColumn());
$columns = [];
$highestRow = $masterDataSheet->getHighestRow();
for ($row = 1; $row <= $highestRow; $row++) {
    $header = $masterDataSheet->getCell('A' . $row)->getValue();
    $columns[chr(64 + $row)] = "masterData!\$B\$$row:\$" . chr(64 + $highestColumn) . "\$$row";
}

// Apply data validation for dropdowns in the importData sheet
foreach ($columns as $col => $formula) {
    echo 'Setting data validation for column ' . $col . PHP_EOL;
    for ($i = 4; $i <= 10; $i++) {
        $validation = $importDataSheet->getCell($col . $i)->getDataValidation();
        $validation->setType(DataValidation::TYPE_LIST);
        $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
        $validation->setAllowBlank(false);
        $validation->setShowInputMessage(true);
        $validation->setShowErrorMessage(true);
        $validation->setShowDropDown(true);
        $validation->setErrorTitle('Input error');
        $validation->setError('Value is not in list');
        $validation->setPromptTitle('Pick from list');
        $validation->setPrompt('Please pick a value from the drop-down list');
        $validation->setFormula1($formula);
    }
}

// Save the spreadsheet
$writer = new Xlsx($spreadsheet);

$writer->save('importData.xlsx');

echo 'Spreadsheet created successfully' . PHP_EOL;
?>