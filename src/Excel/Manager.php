<?php
/*
 * The MIT License
 *
 * Copyright 2017 RobotUsers
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

namespace Robotusers\Excel\Excel;

use Cake\Datasource\EntityInterface;
use Cake\Filesystem\File;
use Cake\ORM\Table;
use DateTimeInterface;
use PHPExcel;
use PHPExcel_Cell;
use PHPExcel_Cell_DataType;
use PHPExcel_IOFactory;
use PHPExcel_Reader_CSV;
use PHPExcel_Reader_IReader;
use PHPExcel_Shared_Date;
use PHPExcel_Worksheet;
use PHPExcel_Worksheet_Row;
use PHPExcel_Writer_IWriter;

/**
 * Description of Manager
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class Manager
{

    /**
     *
     * @param PHPExcel_Worksheet $worksheet
     * @param Table $table
     * @param array $options
     * @return EntityInterface[]
     */
    public function read(PHPExcel_Worksheet $worksheet, Table $table, array $options = [])
    {
        $options += [
            'startRow' => 1,
            'endRow' => null,
            'startColumn' => 'A',
            'endColumn' => null,
            'columnMap' => [],
            'marshallerOptions' => [],
            'saveOptions' => []
        ];

        $columns = $table->getSchema()->columns();
        $primaryKey = $table->getPrimaryKey();

        $rows = $worksheet->getRowIterator($options['startRow'], $options['endRow']);
        $entities = [];
        foreach ($rows as $rowIndex => $row) {
            /* @var $row PHPExcel_Worksheet_Row */

            $data = [
                $primaryKey => $rowIndex
            ];
            $cells = $row->getCellIterator($options['startColumn'], $options['endColumn']);

            $hasData = false;
            foreach ($cells as $cell) {
                /* @var $cell PHPExcel_Cell */

                $column = $cell->getColumn();
                if (isset($options['columnMap'][$column])) {
                    $property = $options['columnMap'][$column];
                } else {
                    $property = $column;
                }

                $value = $cell->getValue();
                if (in_array($property, $columns) && $value !== null) {
                    $data[$property] = $cell->getFormattedValue();
                    $hasData = true;
                }
            }

            if ($hasData) {
                $entity = $table->newEntity($data, $options['marshallerOptions']);
                $table->save($entity, $options['saveOptions']);
                $entities[] = $entity;
            }
        }

        return $entities;
    }

    /**
     *
     * @param PHPExcel_Worksheet $worksheet
     * @param array $options
     * @return PHPExcel_Worksheet
     */
    public function clear(PHPExcel_Worksheet $worksheet, array $options = [])
    {
        $options += [
            'startRow' => 1,
            'endRow' => null,
            'startColumn' => 'A',
            'endColumn' => null,
        ];

        $rows = $worksheet->getRowIterator($options['startRow'], $options['endRow']);
        foreach ($rows as $row) {
            $cells = $row->getCellIterator($options['startColumn'], $options['endColumn']);
            foreach ($cells as $cell) {
                $cell->setValue(null);
            }
        }

        return $worksheet;
    }

    /**
     *
     * @param Table $table
     * @param PHPExcel_Worksheet $worksheet
     * @param array $options
     * @return PHPExcel_Worksheet
     */
    public function write(Table $table, PHPExcel_Worksheet $worksheet, array $options = [])
    {
        $options += [
            'finder' => 'all',
            'finderOptions' => [],
            'propertyMap' => [],
            'startRow' => 1,
            'keepOriginalRows' => false
        ];

        $pk = $table->getPrimaryKey();
        $results = $table->find($options['finder'], $options['finderOptions'])->all();

        $row = $options['startRow'];
        foreach ($results as $result) {
            if ($options['keepOriginalRows']) {
                $row = $result->get($pk);
            }
            $data = $result->toArray();
            unset($data[$pk]);

            foreach ($data as $property => $value) {
                if (isset($options['propertyMap'][$property])) {
                    $column = $options['propertyMap'][$property];
                } else {
                    $column = strtoupper($property);
                }

                $coords = $column . $row;
                $cell = $worksheet->getCell($coords);
                $this->setCellValue($cell, $value);
            }

            $row++;
        }

        return $worksheet;
    }

    /**
     *
     * @param PHPExcel_Cell $cell
     * @param mixed $value
     * @return void
     */
    protected function setCellValue(PHPExcel_Cell $cell, $value)
    {
        if ($value instanceof DateTimeInterface) {
            $value = PHPExcel_Shared_Date::PHPToExcel($value);
        }
        $cell->setValue($value);
        if (is_numeric($value)) {
            $cell->setDataType(PHPExcel_Cell_DataType::TYPE_NUMERIC);
        }
        if (is_bool($value)) {
            $cell->setDataType(PHPExcel_Cell_DataType::TYPE_BOOL);
        }
    }

    /**
     *
     * @param File $file
     * @param array $options
     * @return PHPExcel
     */
    public function getExcel(File $file, array $options = [])
    {
        $reader = $this->getReader($file, $options);

        return $reader->load($file->pwd());
    }

    /**
     *
     * @param File $file
     * @param array $options
     * @return PHPExcel_Reader_IReader
     */
    public function getReader(File $file, array $options = [])
    {
        $reader = PHPExcel_IOFactory::createReaderForFile($file->pwd());

        if ($reader instanceof PHPExcel_Reader_CSV) {
            if (isset($options['delimiter'])) {
                $reader->setDelimiter($options['delimiter']);
            }
        }
        if (isset($options['callback'])) {
            $result = $options['callback']($reader, $file);
            if ($result instanceof PHPExcel_Reader_IReader) {
                $reader = $result;
            }
        }

        return $reader;
    }

    /**
     *
     * @param PHPExcel $excel
     * @param File $file
     * @return PHPExcel_Writer_IWriter
     */
    public function getWriter(PHPExcel $excel, File $file)
    {
        $type = PHPExcel_IOFactory::identify($file->pwd());

        return PHPExcel_IOFactory::createWriter($excel, $type);
    }
}
