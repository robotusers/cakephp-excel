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
use InvalidArgumentException;
use LogicException;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use UnexpectedValueException;

/**
 * Description of Manager
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class Manager
{

    /**
     *
     * @param Worksheet $worksheet
     * @param Table $table
     * @param array $options
     * @return EntityInterface[]
     */
    public function read(Worksheet $worksheet, Table $table, array $options = [])
    {
        $options += [
            'startRow' => 1,
            'endRow' => null,
            'startColumn' => 'A',
            'endColumn' => null,
            'columnMap' => [],
            'marshallerOptions' => [],
            'saveOptions' => [],
            'keepOriginalRows' => false
        ];

        $columns = $table->getSchema()->columns();
        $primaryKey = $table->getPrimaryKey();

        $rows = $worksheet->getRowIterator($options['startRow'], $options['endRow']);
        $entities = [];
        foreach ($rows as $rowIndex => $row) {
            /* @var $row Row */

            $data = [];
            if ($options['keepOriginalRows']) {
                $data[$primaryKey] = $rowIndex;
            }
            $cells = $row->getCellIterator($options['startColumn'], $options['endColumn']);

            $hasData = false;
            foreach ($cells as $cell) {
                /* @var $cell Cell */

                $column = $cell->getColumn();
                $property = $this->resolveKey($options['columnMap'], $column);
                if ($property === null) {
                    continue;
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
     * @param Worksheet $worksheet
     * @param array $options
     * @return Worksheet
     */
    public function clear(Worksheet $worksheet, array $options = [])
    {
        $options += [
            'startRow' => 1,
            'endRow' => null,
            'startColumn' => 'A',
            'endColumn' => null
        ];

        if ($options['startRow'] > $worksheet->getHighestRow()) {
            return $worksheet;
        }

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
     * @param Worksheet $worksheet
     * @param array $options
     * @return Worksheet
     * @throws UnexpectedValueException
     */
    public function write(Table $table, Worksheet $worksheet, array $options = [])
    {
        $options += [
            'finder' => 'all',
            'finderOptions' => [],
            'propertyMap' => [],
            'header' => null,
            'columnCallbacks' => [],
            'startRow' => 1,
            'keepOriginalRows' => false,
            'removePrimaryKey' => true
        ];

        if (is_array($options['header'])) {
            if ($options['startRow'] < 2) {
                $message = 'Option `startRow` must be > 1 if you want to attach header.';
                throw new LogicException($message);
            }

            $this->attachHeader($worksheet, $options['header']);
        }

        $pk = $table->getPrimaryKey();
        $results = $table->find($options['finder'], $options['finderOptions'])->all();

        $row = $options['startRow'];
        foreach ($results as $result) {
            if (is_array($result)) {
                $data = $result;
            } elseif (is_object($result) && method_exists($result, 'toArray')) {
                $data = $result->toArray();
            } else {
                throw new UnexpectedValueException('Cannot convert result to array.');
            }
            if ($options['keepOriginalRows']) {
                $row = $data[$pk];
            }
            if ($options['removePrimaryKey']) {
                unset($data[$pk]);
            }

            foreach ($data as $property => $value) {
                $column = $this->resolveKey($options['propertyMap'], $property);
                if ($column === null) {
                    continue;
                }
                $column = strtoupper($column);

                $coords = $column . $row;
                $cell = $worksheet->getCell($coords);
                $this->setCellValue($cell, $value);
                if (isset($options['columnCallbacks'][$column])) {
                    $callback = $options['columnCallbacks'][$column];
                    $callback($cell, $data);
                }
            }

            $row++;
        }

        return $worksheet;
    }

    /**
     * @param array $keys
     * @param string $key
     * @return string|null
     */
    protected function resolveKey(array $keys, $key)
    {
        $value = true;
        if (isset($keys[$key])) {
            $value = $keys[$key];
        } elseif (isset($keys['*'])) {
            $value = $keys['*'];
        }

        if ($value === true) {
            $value = $key;
        }

        return strlen($value) ? $value : null;
    }

    /**
     *
     * @param Worksheet $worksheet
     * @param array $header
     * @param array $options
     * @return Worksheet
     */
    public function attachHeader(Worksheet $worksheet, array $header, array $options = [])
    {
        $options += [
            'row' => 1,
            'style' => [
                'font' => [
                    'bold' => true
                ]
            ]
        ];

        foreach ($header as $column => $value) {
            $coordinate = strtoupper($column) . $options['row'];
            $worksheet->getCell($coordinate)
                ->setValue($value)
                ->getStyle()
                ->applyFromArray($options['style']);
        }

        return $worksheet;
    }

    /**
     *
     * @param Cell $cell
     * @param mixed $value
     * @return void
     */
    protected function setCellValue(Cell $cell, $value)
    {
        if ($value instanceof DateTimeInterface) {
            $value = Date::PHPToExcel($value->format('U'));
            $cell->getStyle()->getNumberFormat()->setFormatCode('YYYY-MM-DD HH:MM:SS');
        }
        $cell->setValue($value);
        if (is_numeric($value)) {
            $cell->setDataType(DataType::TYPE_NUMERIC);
        }
        if (is_bool($value)) {
            $cell->setDataType(DataType::TYPE_BOOL);
        }
    }

    /**
     *
     * @param File $file
     * @param array $options
     * @return Spreadsheet
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
     * @return IReader
     * @throws InvalidArgumentException
     */
    public function getReader(File $file, array $options = [])
    {
        if (!$file->exists()) {
            $message = sprintf('File %s does not exist.', $file->name());
            throw new InvalidArgumentException($message);
        }

        $reader = IOFactory::createReaderForFile($file->pwd());

        if ($reader instanceof CsvReader) {
            if (isset($options['delimiter'])) {
                $reader->setDelimiter($options['delimiter']);
            }
        }
        if (isset($options['callback'])) {
            trigger_error('Option `callback` has been deprecated. Use `readerCallback` instead', E_USER_DEPRECATED);
        }
        if (isset($options['readerCallback'])) {
            $result = $options['readerCallback']($reader, $file);
            if ($result instanceof IReader) {
                $reader = $result;
            }
        }

        return $reader;
    }

    /**
     *
     * @param Spreadsheet $excel
     * @param File $file
     * @param array $options
     * @return IWriter
     * @throws InvalidArgumentException
     */
    public function getWriter(Spreadsheet $excel, File $file, array $options = [])
    {
        if (!$file->exists()) {
            $message = sprintf('File %s does not exist.', $file->name());
            throw new InvalidArgumentException($message);
        }

        if (isset($options['writerType'])) {
            $type = $options['writerType'];
        } else {
            $type = IOFactory::identify($file->pwd());
        }
        $writer = IOFactory::createWriter($excel, $type);

        if ($writer instanceof CsvWriter) {
            if (isset($options['delimiter'])) {
                $writer->setDelimiter($options['delimiter']);
            }
        }
        if (isset($options['callback'])) {
            trigger_error('Option `callback` has been deprecated. Use `writerCallback` instead', E_USER_DEPRECATED);
        }
        if (isset($options['writerCallback'])) {
            $result = $options['writerCallback']($writer, $file);
            if ($result instanceof IWriter) {
                $writer = $result;
            }
        }

        return $writer;
    }

    /**
     *
     * @param Spreadsheet $excel
     * @param File $file
     * @param array $options
     * @return File
     */
    public function save(Spreadsheet $excel, File $file, array $options = [])
    {
        $writer = $this->getWriter($excel, $file, $options);
        $writer->save($file->pwd());

        return $file;
    }
}
