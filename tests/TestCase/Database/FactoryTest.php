<?php
/*
 * The MIT License
 *
 * Copyright 2017 Robert Pustułka <r.pustulka@robotusers.com>.
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
namespace Robotusers\Excel\Test\TestCase\Database;

use Cake\Database\Connection;
use Cake\Database\Schema\TableSchema;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Robotusers\Excel\Database\Factory;
use Robotusers\Excel\Test\TestCase;

/**
 * Description of FactoryTest
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class FactoryTest extends TestCase
{
    public function testSetNumberFormat()
    {
        $factory = new Factory();

        $factory->setNumberFormat('@', 'string');

        $formats = $factory->getNumberFormatMap();
        $this->assertArrayHasKey('@', $formats);
        $this->assertEquals('string', $formats['@']);
    }

    public function testSetDataType()
    {
        $factory = new Factory();

        $factory->setDataType('n', 'float');

        $types = $factory->getDataTypeMap();
        $this->assertArrayHasKey('n', $types);
        $this->assertEquals('float', $types['n']);
    }

    public function testCreateTable()
    {
        $connection = $this->createMock(Connection::class);
        $schema = $this->createMock(TableSchema::class);
        $queries = [
            'SQL QUERY;'
        ];

        $schema->expects($this->once())
            ->method('createSql')
            ->with($connection)
            ->willReturn($queries);

        $connection->expects($this->once())
            ->method('execute')
            ->with($queries[0]);

        $factory = new Factory();

        $results = $factory->createTable($connection, $schema);
        $this->assertEquals(1, $results);
    }

    public function testGetTableName()
    {
        $excel = $this->createMock(Spreadsheet::class);
        $worksheet = $this->createMock(Worksheet::class);
        $id = 'abcd1234';
        $title = 'Test worksheet 1';

        $worksheet->expects($this->once())
            ->method('getTitle')
            ->willReturn($title);

        $worksheet->expects($this->once())
            ->method('getParent')
            ->willReturn($excel);

        $excel->expects($this->once())
            ->method('getID')
            ->willReturn($id);

        $factory = new Factory();

        $name = $factory->getTableName($worksheet);
        $this->assertEquals('abcd1234_test_worksheet1s', $name);
    }

    public function testPrimaryKey()
    {
        $factory = new Factory();

        $pk = $factory->getPrimaryKey();
        $this->assertEquals('_row', $pk);

        $factory->setPrimaryKey('_id');
        $pk = $factory->getPrimaryKey();
        $this->assertEquals('_id', $pk);
    }

    public function testCreateSchema()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $schema = $factory->createSchema($worksheet);

        $columns = [
            '_row' => 'integer',
            'A' => 'string',
            'B' => 'string',
            'C' => 'string',
            'D' => 'string',
            'E' => 'string',
            'F' => 'string'
        ];
        $this->assertColumns($columns, $schema);
    }

    public function testCreateSchemaStartRowEmpty()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $schema = $factory->createSchema($worksheet, [
            'startRow' => 100
        ]);

        $columns = [
            '_row' => 'integer'
        ];
        $this->assertColumns($columns, $schema);
    }

    public function testCreateSchemaWithHeader()
    {
        $worksheet = $this->loadWorksheet('test.xlsx', 1);
        $this->assertEquals('Sheet Header', $worksheet->getTitle());

        $factory = new Factory();

        $schema = $factory->createSchema($worksheet, [
            'startRow' => 2
        ]);

        $columns = [
            '_row' => 'integer',
            'A' => 'string',
            'B' => 'string',
            'C' => 'string',
            'D' => 'string',
            'E' => 'string',
            'F' => 'string'
        ];
        $this->assertColumns($columns, $schema);
    }

    public function testCreateSchemaTableName()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $name = str_replace('.', '', $worksheet->getParent()->getID()) . '_sheets';

        $schema = $factory->createSchema($worksheet);
        $this->assertEquals($name, $schema->name());

        $schema2 = $factory->createSchema($worksheet, [
            'tableName' => 'excel'
        ]);
        $this->assertEquals('excel', $schema2->name());
    }

    public function testCreateSchemaDefaultType()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $schema = $factory->createSchema($worksheet, [
            'defaultType' => 'integer'
        ]);

        $columns = [
            '_row' => 'integer',
            'A' => 'integer',
            'B' => 'integer',
            'C' => 'integer',
            'D' => 'integer',
            'E' => 'integer',
            'F' => 'integer'
        ];
        $this->assertColumns($columns, $schema);
    }

    public function testCreateSchemaLimitColumns()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $schema = $factory->createSchema($worksheet, [
            'startColumn' => 'C',
            'endColumn' => 'E'
        ]);

        $columns = [
            '_row' => 'integer',
            'C' => 'string',
            'D' => 'string',
            'E' => 'string'
        ];
        $this->assertColumns($columns, $schema);
    }

    public function testCreateSchemaColumnMap()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $schema = $factory->createSchema($worksheet, [
            'columnMap' => [
                'A' => 'columnA',
                'B' => 'columnB'
            ]
        ]);

        $columns = [
            '_row' => 'integer',
            'columnA' => 'string',
            'columnB' => 'string',
            'C' => 'string',
            'D' => 'string',
            'E' => 'string',
            'F' => 'string'
        ];
        $this->assertColumns($columns, $schema);
    }

    public function testCreateSchemaCustomTypes()
    {
        $worksheet = $this->loadWorksheet('test.xlsx');
        $factory = new Factory();

        $schema = $factory->createSchema($worksheet, [
            'numberFormatMap' => [
                NumberFormat::FORMAT_NUMBER => 'integer'
            ],
            'dataTypeMap' => [
                DataType::TYPE_NUMERIC => 'float',
                DataType::TYPE_NULL => false
            ],
            'columnTypeMap' => [
                'D' => 'datetime',
                'E' => 'date',
                'F' => 'time'
            ]
        ]);

        $columns = [
            '_row' => 'integer',
            'A' => 'string',
            'B' => 'integer',
            'C' => 'float',
            'D' => 'datetime',
            'E' => 'date',
            'F' => 'time'
        ];
        $this->assertColumns($columns, $schema);
    }

    protected function assertColumns($columns, $schema)
    {
        $this->assertEquals(array_keys($columns), $schema->columns());
        foreach ($columns as $column => $type) {
            $message = sprintf('Failed asserting that column %s is type of %s', $column, $type);
            $this->assertEquals($type, $schema->getColumnType($column), $message);
        }
    }
}
