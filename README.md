# xlsx2data

PHP Library to convert Microsoft Excel files to xml, yaml or json data 

## Installation

The most convenient way to use the converter is with composer

    composer require bur-gmbh/xlsx2data
    
    
## Usage

To load an Microsoft Excel file just use the constructor and pass the filepath and name of the worksheet:

    $file = 'myExcel.xlsx';
    $sheet = 'Worksheet';

    $converter = new Converter($file, $sheet);
    
If you don't pass a worksheet name the converter will use the active worksheet.
    

### XML

To convert the Excel data to XMl use the `toXml` function.

    $converter->toXml();

### JSON

To convert the Excel data to JSON use the following `toJson` function.

    $converter->toJson()

### YAML

Save the Excel data as Yaml with the `toYaml` command.
    
    $converter->toYaml()

