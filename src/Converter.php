<?php

namespace BUR;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use Symfony\Component\Yaml\Yaml;
use Cocur\Slugify\Slugify;

class Converter
{

    /**
     * @var string
     */
    protected $file;

    protected $data;

    protected $spreadsheet;

    protected $worksheetName;

    protected $worksheet;

    protected $options = [
        'root' => 'data',
        'row' => 'entry',
        'headerRow' => 1,
        'dataRow' => 2
    ];


    protected $headers;

    protected $outputStructure;


    public function __construct(string $file, string $worksheetName = null)
    {
        $realFilePath = realpath($file);

        if ($realFilePath === false) {
            throw new \Exception('File '. $file . ' does not exists');
        }

        $this->worksheetName = $worksheetName;
        $this->file = $realFilePath;

        $this->convert();
    }

    /**
     * @param $key
     * @param null $default
     * @return mixed|null
     */
    protected function getOption($key, $default = null) {

        return $this->options[$key] ?? $default;

    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    protected function convert()
    {

        $reader = new Xlsx();
        $this->spreadsheet = $reader->load($this->file);

        if ($this->worksheetName !== null) {
            $this->worksheet = $this->spreadsheet->getSheetByName($this->worksheetName);
        } else {
            $this->worksheet = $this->spreadsheet->getActiveSheet();
        }

        $this->headers = $this->getHeadersFromExcel($this->getOption('headerRow', 1));
        $this->data = $this->getDataFromExcel($this->getOption('dataRow', 2));

        $this->outputStructure = $this->getOutputStructure();

    }


    protected function getHeadersFromExcel($row) : array
    {

        // get the last index
        $highestColumn = $this->worksheet->getHighestDataColumn();
        $columnAsInteger = Coordinate::columnIndexFromString($highestColumn) - 1;

        $headers = [];

        for ($i = 0; $i <= $columnAsInteger; $i++) {
            $cell = Coordinate::stringFromColumnIndex($i+1). $row;
            $headers[$i] = $this->worksheet->getCell($cell)->getValue();
        }

        return $headers;

    }

    protected function getDataFromExcel($row) : array
    {

        $highestRow = $this->worksheet->getHighestDataRow();
        $highestColumn = $this->worksheet->getHighestDataColumn();
        $columnAsInteger = Coordinate::columnIndexFromString($highestColumn) - 1;

        $data = [];

        for ($row; $row <= $highestRow; $row++) {
            for ($column = 0; $column <= $columnAsInteger; $column++) {
                $cell = Coordinate::stringFromColumnIndex($column+1). $row;
                $data[$row][$column] = $this->worksheet->getCell($cell)->getValue();
            }
        }

        return $data;
    }

    /**
     * @param string $s
     * @return string
     */
    public function normalizeString(string $s) : string {

        $slugify = new Slugify();

        return $slugify->slugify($s);
    }


    /**
     * @return array
     */
    private function getOutputStructure() : array {

        $json = [];

        foreach($this->data as $data) {

            $row = [];

            foreach($this->headers as $key => $attribute) {

                $normalizedHeader = $this->normalizeString($attribute);
                $row[$normalizedHeader] = $data[$key];
            }

            $json[] = $row;
        }

        return $json;
    }

    /**
     * @return string
     */
    public function toYaml() : string
    {

        return Yaml::dump($this->outputStructure, 10);

    }

    /**
     * @return string
     */
    public function toXml() : string
    {

        $xml = new \SimpleXMLElement('<?xml version="1.0" encoding="UTF-8"?><'. $this->getOption('root', 'data') .'/>');

        foreach($this->outputStructure as $row) {

            $entry = $xml->addChild($this->getOption('row', 'entry'));

            foreach($row as $attr => $da) {
                $entry->addChild($attr);
                $entry->{$attr} = $da;
            }
        }

        // for the pretty print
        $dom = dom_import_simplexml($xml)->ownerDocument;
        $dom->formatOutput = true;

        return $dom->saveXML();

    }

    /**
     * @return string
     */
    public function toJson() : string
    {
        return json_encode($this->outputStructure, JSON_PRETTY_PRINT);
    }

    /**
     * @param string $filepath
     * @return bool
     */
    public function saveXml(string $filepath) : bool
    {
        return file_put_contents($filepath, $this->toXml()) !== false;
    }

    /**
     * @param string $filepath
     * @return bool
     */
    public function saveYaml(string $filepath) : bool
    {
        return file_put_contents($filepath, $this->toYaml()) !== false;
    }

    /**
     * @param string $filepath
     * @return bool
     */
    public function saveJson(string $filepath) : bool
    {
        return file_put_contents($filepath, $this->toJson()) !== false;
    }

}