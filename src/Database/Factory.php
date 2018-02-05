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

namespace Robotusers\Excel\Database;

use Cake\Database\Connection;
use Cake\Database\Schema\TableSchema;
use Cake\Utility\Inflector;
use Cake\Utility\Text;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Description of Factory
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class Factory
{

    /**
     *
     * @var array
     */
    protected $dataTypeMap = [
        DataType::TYPE_NULL => false
    ];

    /**
     *
     * @var array
     */
    protected $numberFormatMap = [];

    /**
     *
     * @var string
     */
    protected $primaryKey = '_row';

    /**
     *
     * @param Worksheet $worksheet
     * @param array $options
     * @return TableSchema
     */
    public function createSchema(Worksheet $worksheet, array $options = [])
    {
        $options += [
            'tableName' => $this->getTableName($worksheet),
            'primaryKey' => $this->primaryKey,
            'startRow' => 1,
            'startColumn' => 'A',
            'endColumn' => null,
            'columnMap' => [],
            'dataTypeMap' => $this->dataTypeMap,
            'numberFormatMap' => $this->numberFormatMap,
            'columnTypeMap' => [],
            'defaultType' => [
                'type' => 'string',
                'null' => true
            ]
        ];

        $tableName = $options['tableName'];
        $primaryKey = $options['primaryKey'];
        $schema = new TableSchema($tableName);
        $schema
            ->addColumn($primaryKey, 'integer')
            ->addConstraint('primary', [
                'type' => 'primary',
                'columns' => [$primaryKey]
            ]);

        $row = new Row($worksheet, $options['startRow']);
        $cells = $row->getCellIterator($options['startColumn'], $options['endColumn']);
        foreach ($cells as $cell) {
            /* @var $cell Cell */

            $type = $this->discoverType($cell, $options);

            if ($type !== false) {
                $column = $cell->getColumn();
                if (isset($options['columnMap'][$column])) {
                    $column = $options['columnMap'][$column];
                }
                $schema->addColumn($column, $type);
            }
        }

        return $schema;
    }

    /**
     *
     * @param Cell $cell
     * @param array $options
     * @return array|string|bool
     */
    protected function discoverType(Cell $cell, array $options)
    {
        $format = $cell->getStyle()->getNumberFormat()->getFormatCode();
        $dataType = $cell->getDataType();
        $column = $cell->getColumn();

        if (isset($options['columnTypeMap'][$column])) {
            $type = $options['columnTypeMap'][$column];
        } elseif (isset($options['numberFormatMap'][$format])) {
            $type = $options['numberFormatMap'][$format];
        } elseif (isset($options['dataTypeMap'][$dataType])) {
            $type = $options['dataTypeMap'][$dataType];
        } else {
            $type = $options['defaultType'];
        }

        return $type;
    }

    /**
     *
     * @param Connection $connection
     * @param TableSchema $schema
     * @return int
     */
    public function createTable(Connection $connection, TableSchema $schema)
    {
        $queries = $schema->createSql($connection);

        foreach ($queries as $sql) {
            $connection->execute($sql);
        }

        return count($queries);
    }

    /**
     *
     * @param Worksheet $worksheet
     * @return string
     */
    public function getTableName(Worksheet $worksheet)
    {
        $excel = $worksheet->getParent();
        $title = $excel->getID() . ' ' . $worksheet->getTitle();

        $slug = Text::slug($title, [
            'replacement' => '_'
        ]);
        $camelized = Inflector::camelize($slug);
        $name = Inflector::tableize($camelized);

        return $name;
    }

    /**
     *
     * @return array
     */
    public function getDataTypeMap()
    {
        return $this->dataTypeMap;
    }

    /**
     *
     * @return array
     */
    public function getNumberFormatMap()
    {
        return $this->numberFormatMap;
    }

    /**
     *
     * @param string $type
     * @param string|array|bool $column
     * @return $this
     */
    public function setDataType($type, $column)
    {
        $this->dataTypeMap[$type] = $column;

        return $this;
    }

    /**
     *
     * @param string $format
     * @param string|array|bool $column
     * @return $this
     */
    public function setNumberFormat($format, $column)
    {
        $this->numberFormatMap[$format] = $column;

        return $this;
    }

    /**
     *
     * @return string
     */
    public function getPrimaryKey()
    {
        return $this->primaryKey;
    }

    /**
     *
     * @param string $primaryKey
     * @return $this
     */
    public function setPrimaryKey($primaryKey)
    {
        $this->primaryKey = $primaryKey;

        return $this;
    }
}
