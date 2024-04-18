<?php
set_time_limit(300);
require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\IOFactory;

class ImportToExcel
{

  private $targetFile;
  private $folder;
  private $alias;

  public function __construct($targetFile, $folderToScan)
  {
    $this->targetFile = $targetFile;
    $this->folder = __DIR__ . "/" . $folderToScan;
    $this->alias = substr($this->targetFile, 0, 3);
    $this->scanFolder();
  }

  public function scanFolder()
  {
    $files = scandir($this->folder);
    $files = array_diff($files, ['.', '..']);
    $filtered_files = array_filter($files, [$this, "filter"]);

    $bounces = [];
    $emails = [];
    foreach ($filtered_files as $key => $file) {
      # code...
      if (preg_match("/bounce/", $file)) {
        if (!empty($this->getBounces($file))) {
          $bounces = array_merge($this->getBounces($file), $bounces);
        }
      } else {
        if (!empty($this->getDailyActivity($file))) {
          $emails = array_merge($this->getDailyActivity($file), $emails);
        }
      }
    }

    // if (count($bounces) > 0) {
    //   $this->insertBounces($bounces);
    // }

    // if (count($emails) > 0) {
    //   $this->insertEmails($emails);
    // }

    // unset($bounces);
    // unset($emails);
  }

  public function filter($file_name)
  {
    $full_path = $this->folder . '/' . $file_name;
    if (pathinfo($full_path, PATHINFO_EXTENSION) == "csv" && filesize($full_path) > 100) {
      return  true;
    }
    return false;
  }

  public function getDailyActivity($file)
  {
    echo "Getting Emails from $file ...<br>\n";
    $excelFile = $this->folder . '/' . $file; // Replace with the path to your Excel file
    $spreadsheet = IOFactory::load($excelFile);

    $worksheet = $spreadsheet->getSheet(0);
    $worksheet->removeColumnByIndex(1);
    // $worksheet->removeColumnByIndex(3);
    $worksheet->removeColumnByIndex(4);
    $worksheet->removeColumnByIndex(7);
    $worksheet->removeColumnByIndex(8);
    $worksheet->removeColumnByIndex(8);
    $worksheet->removeColumnByIndex(8);
    $worksheet->removeColumnByIndex(8);
    $worksheet->removeColumnByIndex(8);

    $highestRow = $worksheet->getHighestRow();
    $highestColumn = $worksheet->getHighestColumn();

    $emails = [];
    for ($row = 2; $row <= $highestRow; $row++) {
      $tmpRow = [];
      for ($col = 'A'; $col <= $highestColumn; $col++) {
        $cellValue = $worksheet->getCell($col . $row)->getValue();
        $tmpRow[$col] = ($col == "C" && $tmpRow["B"] == "grace.mendoza@robinsonsland.com") ? "lea.tan@robinsonsland.com" : $cellValue;
      }
      $emails[] = $tmpRow;
    }
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);
    $emails = array_filter($emails, function ($value) {
      return !empty($value["F"]) && strpos($value["F"], $this->alias) !== false;
    });
    return $emails;
  }

  public function insertEmails($emails)
  {
    echo "Inserting Emails to " . $this->targetFile . " ...<br>\n";
    $targetFile = $this->targetFile;

    $reader = IOFactory::createReader("Xlsx");
    $spreadsheet = $reader->load($targetFile);

    $writer = IOFactory::createWriter($spreadsheet, "Xlsx");
    $worksheet = $spreadsheet->setActiveSheetIndex(2);
    # Write in Bound
    foreach ($emails as $key => $email) {
      $index = $worksheet->getHighestDataRow() + 1;
      for ($col = 'A'; $col <= "G"; $col++) {
        $worksheet->setCellValue($col . $index, $emails[$key][$col]);
      }
    }

    $writer->save($targetFile);
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);
    unset($emails);
  }

  public function getBounces($file)
  {
    echo "Getting Bounces from $file ...<br>\n";
    $excelFile = $this->folder . '/' . $file; // Replace with the path to your Excel file
    $spreadsheet = IOFactory::load($excelFile);

    $worksheet = $spreadsheet->getSheet(0);
    $worksheet->removeColumnByIndex(2);
    $worksheet->removeColumnByIndex(3);
    $worksheet->removeColumnByIndex(3);
    $worksheet->removeColumnByIndex(3);
    $worksheet->removeColumnByIndex(3);
    $worksheet->removeColumnByIndex(3);
    $worksheet->removeColumnByIndex(9);
    $worksheet->removeColumnByIndex(9);
    $worksheet->removeColumnByIndex(9);

    $highestRow = $worksheet->getHighestRow();
    $highestColumn = $worksheet->getHighestColumn();

    $bounces = [];
    for ($row = 2; $row <= $highestRow; $row++) {
      $tmpRow = [];
      for ($col = 'A'; $col <= $highestColumn; $col++) {
        $cellValue = $worksheet->getCell($col . $row)->getValue();
        $tmpRow[$col] = $cellValue;
      }
      $bounces[] = $tmpRow;
    }
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);

    $bounces = array_filter($bounces, function ($value) {
      return !empty($value["I"]) && strpos($value["I"], $this->alias) !== false;
    });

    return $bounces;
  }

  public function insertBounces($bounces)
  {
    echo "Inserting Bounces to " . $this->targetFile . " ...<br>\n";
    $targetFile = $this->targetFile;

    $reader = IOFactory::createReader("Xlsx");
    $spreadsheet = $reader->load($targetFile);

    $writer = IOFactory::createWriter($spreadsheet, "Xlsx");
    $worksheet = $spreadsheet->setActiveSheetIndex(1);

    # Write in Bound
    foreach ($bounces as $key => $bounce) {
      $index = $worksheet->getHighestDataRow() + 1;
      for ($col = 'A'; $col <= "I"; $col++) {
        $worksheet->setCellValue($col . $index, $bounces[$key][$col]);
      }
    }

    $writer->save($targetFile);
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);
    unset($bounces);
  }
}

// APC
new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-28");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-29");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-30");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-31");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-1");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-2");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-3");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-4");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-5");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-6");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-7");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-8");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-9");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-10");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-11");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-12");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-13");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-14");
// new ImportToExcel("APC-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-15");

// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-28");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-29");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-30");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-03-31");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-1");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-2");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-3");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-4");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-5");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-6");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-7");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-8");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-9");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-10");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-11");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-12");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-13");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-14");
// new ImportToExcel("CGR-MAIL DELIVERY REPORT_APRIL 2024.xlsx", "2024-04-15");

// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-28");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-29");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-30");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-31");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-1");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-2");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-3");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-4");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-5");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-6");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-7");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-8");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-9");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-10");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-11");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-12");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-13");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-14");
// new ImportToExcel("EET-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-15");

// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-28");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-29");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-30");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-31");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-1");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-2");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-3");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-4");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-5");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-6");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-7");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-8");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-9");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-10");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-11");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-12");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-13");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-14");
// new ImportToExcel("EOG-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-15");

// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-28");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-29");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-30");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-03-31");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-1");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-2");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-3");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-4");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-5");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-6");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-7");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-8");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-9");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-10");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-11");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-12");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-13");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-14");
// new ImportToExcel("FAP-MAIL DELIVERY REPORT_MARCH 2024.xlsx", "2024-04-15");

// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-03-28");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-03-29");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-03-30");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-03-31");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-1");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-2");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-3");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-4");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-5");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-6");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-7");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-8");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-9");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-10");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-11");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-12");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-13");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-14");
// new ImportToExcel("MPR-MAIL DELIVERY REPORT_MARCH.xlsx", "2024-04-15");

// // Load the Excel file