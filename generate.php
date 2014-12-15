<?php
/**
 * Create test cases from an excel file.
 *
 * Usage:
 * php generate.php
 *
 */

// Include autoloader for dependencies
include 'vendor/autoload.php';

// Just for consistency
date_default_timezone_set("Europe/Istanbul");

// Set arguments to $_GET variable.
parse_str(implode('&', array_slice($argv, 1)), $_GET);

// Get filename
$fileNameTestCases = './test-cases.xlsx';
$fileNameTestSpecifications = './test-specifications.xlsx';

// Read test cases
try {
    $fileType = PHPExcel_IOFactory::identify($fileNameTestCases);
    $objReader = PHPExcel_IOFactory::createReader($fileType);
    $testCasesPHPExcel = $objReader->load($fileNameTestCases);
} catch (Exception $e) {
    die('Error loading file "' . pathinfo($fileNameTestCases, PATHINFO_BASENAME) . '": ' . $e->getMessage());
}

// Read test specifications
try {
    $fileType = PHPExcel_IOFactory::identify($fileNameTestCases);
    $objReader = PHPExcel_IOFactory::createReader($fileType);
    $testSpecificationsPHPExcel = $objReader->load($fileNameTestSpecifications);
} catch (Exception $e) {
    die('Error loading file "' . pathinfo($fileNameTestSpecifications, PATHINFO_BASENAME) . '": ' . $e->getMessage());
}

// Get first sheet. We do not support multiple sheets
$sheetTestCases = $testCasesPHPExcel->getSheet(0);
$sheetSpecificationsCases = $testSpecificationsPHPExcel->getSheet(0);

// Read whole data of two files.
$testCases = $sheetTestCases->rangeToArray('A2:I145', null, true, false);
$testSpecifications = $sheetSpecificationsCases->rangeToArray('A2:G58', null, true, false);

// Create test cases
$i = 0;
foreach ($testCases as $testCase) {

    $i++;

    // Set filename
    $tcfilename = str_replace('.', '/', $testCase[0]) . ".docx";

    // Create directory of file if not exists
    if (!is_dir('./export/' . dirname($tcfilename))) {
        mkdir('./export/' . dirname($tcfilename), 0755, true);
    }

    // Give feedback to user
    echo "Creating {$i}. test case: {$tcfilename}." . PHP_EOL;

    $styleCell = array('valign' => 'top');
    $fontStyle = array('align' => 'left', 'size' => '12');

    $word = new \PhpOffice\PhpWord\PhpWord();

    $section = $word->addSection();

    $table = $section->addTable('table1');
    $table->setWidth(5000);

    $tabledata = array(
        array('TC ID', $testCase[0]),
        array('Purpose', $testCase[1]),
        array('Requirements', $testCase[2]),
        array('Priority', $testCase[3]),
        array('Est. Time Needed', $testCase[4]),
        array('Dependency', $testCase[5]),
        array('Setup', $testCase[6]),
        array('Procedure', $testCase[7]),
        array('Clean Up', $testCase[8])
    );

    foreach ($tabledata as $data) {
        $table->addRow();
        $table->addCell(800, $styleCell)->addText($data[0], $fontStyle);
        $table->addCell(200, $styleCell)->addText(':', $fontStyle);
        $contentCell = $table->addCell(4000, $styleCell);
        foreach (explode("\n", $data[1]) as $line) {
            $contentCell->addText($line, $fontStyle);
        }
        $contentCell->addText(PHP_EOL, $fontStyle);
    }

    $writer = \PhpOffice\PhpWord\IOFactory::createWriter($word);
    $writer->save("./export/{$tcfilename}");

};

echo "Creating \"Test Design Specifications\"." . PHP_EOL;


$word = new \PhpOffice\PhpWord\PhpWord();
$word->addParagraphStyle('pStyle', array('align' => 'center', 'spaceAfter' => 100));
$word->addTitleStyle(2, array('bold' => true, 'size' => '15'), array('spaceAfter' => 240, 'spaceBefore' => 240));
$word->addTitleStyle(3, array('bold' => true, 'size' => '13'), array('spaceAfter' => 120, 'spaceBefore' => 120));

// Create test design specifications
$styleCell = array('valign' => 'top');
$fontStyle = array('align' => 'left', 'size' => '12');

$styleTable = array('borderSize' => 6, 'borderColor' => '006699', 'cellMargin' => 80);
$styleFirstRow = array('borderBottomSize' => 18, 'borderBottomColor' => '0000FF', 'bgColor' => '66BBFF');
$word->addTableStyle('table1', $styleTable, $styleFirstRow);

$section = $word->addSection();

$i = 0;
foreach ($testSpecifications as $testSpecification) {
    $i++;

    $section->addTitle('18.' . $i . '. ' . $testSpecification[1], 2);
    $section->addTitle('18.' . $i . '.1. Subfeatures to be tested', 3);
    $section->addText($testSpecification[2]);
    $section->addTitle('18.' . $i . '.2. Subfeatures not to be tested', 3);
    $section->addText($testSpecification[3]);
    $section->addTitle('18.' . $i . '.3. Approach', 3);
    $section->addText($testSpecification[4]);
    $section->addTitle('18.' . $i . '.4. Item Pass/Fail Criteria', 3);
    $section->addText($testSpecification[5]);
    $section->addTitle('18.' . $i . '.5. Environmental Needs', 3);
    $section->addText($testSpecification[6]);
    $section->addTitle('18.' . $i . '.6. Test Cases', 3);

    $table = $section->addTable('table1');
    $table->setWidth(5000);

    $table->addRow();
    $table->addCell(100, $styleCell)->addText('TC ID', array('align' => 'left', 'size' => '9', 'bold' => true));
    $table->addCell(100, $styleCell)->addText('Requirements', array('align' => 'left', 'size' => '9', 'bold' => true));
    $table->addCell(100, $styleCell)->addText('Priority', array('align' => 'left', 'size' => '9', 'bold' => true));
    $table->addCell(4700, $styleCell)->addText('Scenario Description', array('align' => 'left', 'size' => '9', 'bold' => true));

    foreach ($testCases as $testCase) {
        if (trim($testSpecification[0]) == trim($testCase[2])) {
            $rowdata = array(
                $testCase[0],
                $testCase[2],
                $testCase[3],
                $testCase[1],
            );
            $table->addRow();
            foreach ($rowdata as $data) {
                $contentCell = $table->addCell(null, $styleCell);
                foreach (explode("\n", $data) as $line) {
                    $contentCell->addText($line, array('align' => 'left', 'size' => '8'));
                }
            }
        }
    }

    $section->addPageBreak();

};

$writer = \PhpOffice\PhpWord\IOFactory::createWriter($word);
$writer->save("./export/Test Design Specifications.docx");

echo "Export completed succefully. " . PHP_EOL;