# CakePHP Excel plugin 

[![Software License](https://img.shields.io/badge/license-MIT-brightgreen.svg)](LICENSE)
[![Build Status](https://travis-ci.org/robotusers/cakephp-excel.svg?branch=master)](https://travis-ci.org/robotusers/cakephp-excel)

CakePHP Excel plugin allows for spreadsheet files manipulation with the power of CakePHP ORM.
This plugin is build using [PHPExcel](https://github.com/PHPOffice/PHPExcel) library and can work with multiple types of spreadsheet files (excel, csv etc).

## Installation

```
composer require robotusers/cakephp-excel
```

## Using the plugin

Excel plugin lets you manipulate spreadsheet files multiple ways. The simplest use case is to load your spreadhseet data into CakePHP ORM table.

For example we are loading an excel file that contains some record data.

|   | A             | B                  | C    |
|:--|:------------: |:------------------:| :---:|
| 1 | Led Zeppelin	| Led Zeppelin II    | 1969 |
| 2 | Deep Purple   | Machine Head       | 1972 |
| 3 | Pink Floyd    | Wish You Were Here | 1975 |

```php
use Robotusers/Excel/Registry;

$registry = Registry::instance();
$table = $registry->get('path/to/records.xlsx', 'Albums');
```

Spreadsheet data is now loaded into CakePHP ORM table.


```php
$row = $table->find()->first()->toArray();

//this is how a simple row looks like:
[
    '_row' => 1,
    'A' => 'Led Zeppelin',
    'B' => 'Led Zeppelin II',
    'C' => '1969'
]
```

Each column is represented as a property. Values are `string` by default. Row number is represented as `_row` property which is also a primary key of a table.

You may also map columns to custom properties and types.

```php
use Robotusers/Excel/Registry;

$registry = Registry::instance();
$table = $registry->get('path/to/records.xlsx', 'Albums', [
    'primaryKey' => 'id',
    'columnMap' => [
        'A' => 'band',
        'B' => 'album',
        'C' => 'year'
    ],
    'columnTypeMap' => [
        'C' => 'date'
    ]
]);
```

Spreadsheet data is now loaded into CakePHP ORM with custom properties and types.


```php
$row = $table->find()->first()->toArray();

//this is how a simple row looks like:
[
    'id' => 1,
    'band' => 'Led Zeppelin',
    'album' => 'Led Zeppelin II',
    'year' => object(Cake\I18n\Date) {
        'time' => '1969-01-01T00:00:00+00:00',
        'timezone' => 'UTC'
    }
]
```

You may want to manipulate some data and write it back to excel file. This is also possible.

```php
$row = $table->newEntity([
    'band' => 'Genesis',
    'album' => 'Selling England by the Pound',
    'year' => '1973'
]);
$table->save($row);
```

Now the new record is saved, but excel file has not been updated yet. You have to call `writeExcel()` method:

```php
$table->writeExcel();
```

## Behavior

This plugin provides a behavior which could be added to any table.

```php
//AlbumsTable.php

public function initialize()
{
    $this->addBehavior('Robotusers/Excel.Excel', [
        'columnMap' => [
            'A' => 'band',
            'B' => 'album',
            'C' => 'year'
        ]
    ]);
}
```

If you want to load data into your table you have to set a worksheet instance.

```php
use Cake\Filesystem\File;

$file = new File('path/to/file.xls');
$excel = $table->getManager()->getExcel($file); // PHPExcel instance
$worksheet = $excel->getActiveSheet(); // PHPExcel_Worksheet instance

$table->setWorksheet($worksheet)->readExcel();
```

Now your table is populated with excel data.

If you want to write your data back to excel file you have to set a file.

```php
$table->setFile($file)->writeExcel();
```

## Working with different tables

It is also possible to load data into any table.

```php
use Robotusers\Excel\Excel\Manager;

$table = TableRegistry::get('SomeTable');
$manager = new Manager();

$file = new File('file.xlsx');
$excel = $manager->getExcel($file);
$worksheet = $excel->getActiveSheet();

$manager->read($worksheet, $table, [
    'columnMap' => [
        'A' => 'band',
        'B' => 'album',
        'C' => 'year'
    ]
]);

//manipulate your data...

//here you have to tell where properties should be placed
$manager->write($table, $worksheet, [
    'propertyMap' => [
        'band' => 'A',
        'album' => 'B',
        'year' => 'C'
    ]
]);
//to actually save the file you have to call save()
$writer = $manager->save($excel, $file);
```
