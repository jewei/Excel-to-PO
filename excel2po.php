<?php
/**
 * Create PO file from Excel with PHPExcel
 */
define('INPUT_FILE_NAME', 'Translations_FR.xlsx');

// Include path
set_include_path(get_include_path() . PATH_SEPARATOR . './Classes/');

// Use PHPExcel_IOFactory
include 'PHPExcel/IOFactory.php';

// Identify a reader to use
$inputFileType = PHPExcel_IOFactory::identify(INPUT_FILE_NAME);

// Create a new Reader of the type defined in $inputFileType
$objReader = PHPExcel_IOFactory::createReader($inputFileType);

// Load $inputFileName to a PHPExcel Object
$objPHPExcel = $objReader->load(INPUT_FILE_NAME);

// Load all sheets
$engine = new PHPExcellent;
$engine->loadExcel($objPHPExcel->getAllSheets());
$data = $engine->toPOFormat();

$duplicates = $engine->getDuplicates();
$missings = $engine->getMissings();

// Out put summary to the browser
ob_start(); ?>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    </head>
    <body>
        <?php if(!empty($duplicates)) { echo htmlList($duplicates, 'Duplicates (' . count($duplicates) . ')'); } ?>
        <?php if(!empty($missings)) { echo htmlList($missings, 'Missings (' . count($missings) . ')'); } ?>
        <h2>Result (<?php echo $engine->getTranslatedCount(); ?>)</h2>
        <pre><?php echo htmlspecialchars($data); ?></pre>
    </body>
</html>
<?php
$content = ob_get_clean();

echo $content;

// Write to po file with proper header
$outputFileName = explode('.', INPUT_FILE_NAME);
$outputFileName = $outputFileName[0] . '.po';
$data = 'msgid ""
msgstr ""
"MIME-Version: 1.0\n"
"Content-Type: text/plain; charset=UTF-8\n"
"Content-Transfer-Encoding: 8bit\n"' . "\n\n" . $data;
file_put_contents($outputFileName, $data);
exit;

function htmlList($list = array(), $header = '') {
    ob_start(); ?>
<h2><?php echo htmlspecialchars($header); ?></h2>
<ul>
<?php foreach ($list as $key => $values) { ?>
    <li>
        <?php echo htmlspecialchars($key); ?>
        <ul>
            <?php foreach ($values as $value) { ?>
            <li><?php echo htmlspecialchars($value); ?></li>
            <?php } ?>
        </ul>
    </li>
<?php } //end foreach ?>
</ul>
<?php
    $content = ob_get_clean();
    return $content;
}

class PHPExcellent
{
    public $translatedCount = 0;
    protected $duplicates = array();
    protected $missings = array();
    protected $excel = '';

    public function getTranslatedCount() {
        return $this->translatedCount;
    }

    public function getDuplicates() {
        return $this->duplicates;
    }

    public function getMissings() {
        return $this->missings;
    }

    public function loadExcel($file) {
        $this->excel = $file;
    }

    public function escape($char) {
        return str_ireplace('"', '\"', $char);
    }

    public function construct($originalString, $translatedString) {
        return 'msgid "' . $this->escape($originalString) . '"' . "\n" . 'msgstr "' . $this->escape($translatedString) . '"';
    }

    public function toPOFormat() {
        $return = '';
        $pairs = array();

        // Loop all the workbooks
        foreach ($this->excel as $sheet) {
            $sheetData = $sheet->toArray(null,true,true,true);

            // Remove the first line. Usually it's the sheet header row "English: Translation"
            array_shift($sheetData);

            // Get the first 2 column in each sheet
            foreach ($sheetData as $pairValue) {
                if( isset($pairValue['A']) && isset($pairValue['B']) ) {

                    // Checking if there's any duplicated entries
                    if(!isset($pairs[$pairValue['A']])) {
                        $pairs[$pairValue['A']] = $pairValue['B'];
                    } else {
                        $this->duplicates[$pairValue['A']][] = $pairValue['B'];
                        continue;
                    }

                    // Only process if both cells has value
                    if( strlen(trim($pairValue['A'])) > 0 && strlen(trim($pairValue['B'])) > 0 ) {
                        $return .= $this->construct($pairValue['A'], $pairValue['B']) . "\n\n";
                        $this->translatedCount++;
                    } else {
                        if( (strlen($pairValue['A']) + strlen($pairValue['A'])) > 0 ) {
                            $this->missings[] = array($pairValue['A'], $pairValue['B']);
                        }
                    }
                }
            }
        }
        return $return;
    }
}
