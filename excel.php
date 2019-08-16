
<?php
// COMPOSER
require 'vendor/autoload.php';

// PHPSPREADSHEET
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// FILE NAME
$fileName = "reader.xlsx";

// READER SETTINGS
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$reader->setReadDataOnly(true);
$reader->setLoadAllSheets();

// LOADING THE SPREADSHEET
$spreadsheet = $reader->load($fileName);
$worksheet = $spreadsheet->getActiveSheet();

// SQLi SETTINGS
$conn = mysqli_connect('localhost', 'root', '');
if (!$conn){
    die('Could not connect: ' . mysql_error());
}

$db_selected = mysqli_select_db($conn, 'excel');
if (!$db_selected){
    die ("Can't use table 'excel' : " . mysqli_error($conn));
};

// ARRAYS TO BE USED IN LOOP() FUNCTION
$ids = [];
$names = [];
$ages = [];

// LOOP FUNCTION
function loop() {
  global $worksheet;
  $highestRow = $worksheet->getHighestRow();
  echo $highestRow . PHP_EOL;
  $highestColumn = $worksheet->getHighestColumn();
  echo $highestColumn . PHP_EOL;
  $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
  echo $highestColumnIndex . PHP_EOL;

  // NEED TO BE GLOBAL IN ORDER TO UPDATE OUTSIDE FUNCTION
  // Not too sure about this one chief. Try creating a getter & setter for these guys. 
  // If php supports classes, you can do it that way too.
  // ie, globals are 99.9% times not the answer.
  global $ids;
  global $names;
  global $ages;

  for ($row = 2; $row <= $highestRow; ++$row) { // START FROM ROW 2 TO AVOID TEMPLATE
    for ($col = 1; $col <= $highestColumnIndex; ++$col) { // START FROM COLUMN A
        $value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
        // SORTING INTO CORRECT ARRAYS
        if ($col == 1){
          $ids[] = $value;
        };
        if ($col == 2){
          $names[] = $value;
        };
        if ($col == 3){
          $ages[] = $value;
        };
    }
  }
};

// TESTING IF FILE FITS TEMPLATE AND RUNNING LOOP FUNCTION IF TRUE
$idTemplate = $worksheet->getCell("A1");
$nameTemplate = $worksheet->getCell("B1");
$ageTemplate = $worksheet->getCell("C1");

# add loop to conditional. You'll need to ask it to return a bool
if ($idTemplate == "ID" && $nameTemplate == "NAME" && $ageTemplate == "AGE" && loop()){
  print_r($ids) . "\n";
  print_r($names) . "\n";
  print_r($ages) . "\n";
}
else{
  echo "INVALID TEMPLATE";
};

$count = 0;

foreach ($ids as $id){
  $id = intval($id);

  $name = $names[$count];

  $age = $ages[$count];
  $age = intval($age);

  $sql = "INSERT INTO main (id, `name`, age)
          VALUES ($id, '$name', $age)";
  $result = mysqli_query($conn, $sql) or die(mysqli_error($conn));

  $count += 1;
  };

?>
