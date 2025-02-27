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
    ['Kelas', '12', '13'],
    ['Kurikulum', 'K13', 'KTSP'],
    ['Mata Pelajaran', 'Matematika', 'Bahasa Indonesia'],
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
    ['Kelas', 'Kurikulum', 'Mata Pelajaran', 'Nama Mata pelajaran', 'Semester', 'ID Silabus', 'Bab', 'Pokok Pembahasan'],
];

// Set the values for the dropdowns in the importData sheet
$dropdowns = [
    ['Kelas', 'masterData!$B$1:$C$1'],
    ['Kurikulum', 'masterData!$B$2:$C$2'],
    ['Mata Pelajaran', 'masterData!$B$3:$C$3'],
    ['Semester', 'masterData!$B$4:$C$4'],
];

// Set the data and dropdowns in the importData sheet
$row = 1;
foreach ($data as $rowData) {
    $col = 'A';
    foreach ($rowData as $value) {
        $importDataSheet->setCellValue($col . $row, $value);
        $col++;
    }
    $row++;
}

$row = 1;
foreach ($dropdowns as $dropdown) {
    $validation = $importDataSheet->getCell('B' . $row)->getDataValidation();
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
    $validation->setFormula1($dropdown[1]);
    $row++;
}

// Save the spreadsheet
$writer = new Xlsx($spreadsheet);

$writer->save('importData.xlsx');

echo 'Spreadsheet created successfully' . PHP_EOL;

?>