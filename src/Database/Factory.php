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
use PHPExcel_Cell;
use PHPExcel_Cell_DataType;
use PHPExcel_Style_NumberFormat;
use PHPExcel_Worksheet;
use PHPExcel_Worksheet_Row;

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
    protected $typeMap = [
        PHPExcel_Cell_DataType::TYPE_NUMERIC => [
            'type' => 'float',
            'null' => true
        ],
        PHPExcel_Cell_DataType::TYPE_BOOL => [
            'type' => 'boolean',
            'null' => true
        ],
        PHPExcel_Cell_DataType::TYPE_NULL => false
    ];

    /**
     *
     * @var array
     */
    protected $formatMap = [
        PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME => [
            'type' => 'datetime',
            'null' => true
        ]
    ];

    /**
     *
     * @param PHPExcel_Worksheet $worksheet
     * @param array $options
     * @return TableSchema
     */
    public function createSchema(PHPExcel_Worksheet $worksheet, array $options = [])
    {
        $options += [
            'table' => $this->getTableName($worksheet),
            'startRow' => 1,
            'startColumn' => 'A',
            'endColumn' => null,
            'typeMap' => $this->typeMap,
            'formatMap' => $this->formatMap,
            'defaultType' => [
                'type' => 'string',
                'null' => true
            ]
        ];

        $name = $options['table'];
        $schema = new TableSchema($name);
        $schema
            ->addColumn('_row', 'integer')
            ->addConstraint('primary', [
                'type' => 'primary',
                'columns' => ['_row']
            ]);

        $row = new PHPExcel_Worksheet_Row($worksheet, $options['startRow']);
        $cells = $row->getCellIterator($options['startColumn'], $options['endColumn']);
        foreach ($cells as $cell) {
            /* @var $cell PHPExcel_Cell */

            $name = strtolower($cell->getColumn());
            $format = $cell->getStyle()->getNumberFormat()->getFormatCode();
            $type = $cell->getDataType();

            if (isset($options['formatMap'][$format])) {
                $type = $options['formatMap'][$format];
            } elseif (isset($options['typeMap'][$type])) {
                $type = $options['typeMap'][$type];
            } else {
                $type = $options['defaultType'];
            }

            if ($type !== false) {
                $schema->addColumn($name, $type);
            }
        }

        return $schema;
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
     * @param PHPExcel_Worksheet $worksheet
     * @return string
     */
    public function getTableName(PHPExcel_Worksheet $worksheet)
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
    public function getTypeMap()
    {
        return $this->typeMap;
    }

    /**
     *
     * @return array
     */
    public function getFormatMap()
    {
        return $this->formatMap;
    }

    /**
     *
     * @param string $type
     * @param string|array|false $column
     * @return $this
     */
    public function setType($type, $column)
    {
        $this->typeMap[$type] = $column;

        return $this;
    }

    /**
     *
     * @param string $format
     * @param string|array|false $column
     * @return $this
     */
    public function setFormat($format, $column)
    {
        $this->formatMap[$format] = $column;

        return $this;
    }
}