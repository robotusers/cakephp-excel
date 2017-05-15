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
namespace Robotusers\Excel\Test\TestCase\Excel;

use Cake\Chronos\Chronos;
use Cake\Chronos\Date;
use Cake\Filesystem\File;
use Cake\ORM\TableRegistry;
use InvalidArgumentException;
use PHPExcel;
use PHPExcel_Reader_CSV;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Writer_CSV;
use PHPExcel_Writer_Excel2007;
use PHPExcel_Writer_IWriter;
use Robotusers\Excel\Excel\Manager;
use Robotusers\Excel\Test\TestCase;
use UnexpectedValueException;

/**
 * Description of ManagerTest
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class ManagerTest extends TestCase
{
    public $fixtures = [
        'plugin.Robotusers/Excel.regular_columns',
        'plugin.Robotusers/Excel.mapped_columns'
    ];

    public function testGetReaderXlsx()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $reader = $manager->getReader($file);
        $this->assertInstanceOf(PHPExcel_Reader_Excel2007::class, $reader);
    }

    /**
     * @expectedException InvalidArgumentException
     * @expectedExceptionMessage File foo does not exist.
     */
    public function testGetReaderMissingFile()
    {
        $manager = new Manager();
        $file = $this->createMock(File::class);
        $file->method('exists')
            ->willReturn(false);
        $file->method('name')
            ->willReturn('foo');

        $manager->getReader($file);
    }

    /**
     * @expectedException InvalidArgumentException
     * @expectedExceptionMessage File foo does not exist.
     */
    public function testGetWriterMissingFile()
    {
        $manager = new Manager();
        $file = $this->createMock(File::class);
        $excel = $this->createMock(PHPExcel::class);
        $file->method('exists')
            ->willReturn(false);
        $file->method('name')
            ->willReturn('foo');

        $manager->getWriter($excel, $file);
    }

    public function testGetReaderCsv()
    {
        $manager = new Manager();
        $file = $this->getFile('test.csv');

        $reader = $manager->getReader($file, [
            'delimiter' => 'FOO'
        ]);
        $this->assertInstanceOf(PHPExcel_Reader_CSV::class, $reader);
        $this->assertEquals('FOO', $reader->getDelimiter());
    }

    public function testGetReaderCallback()
    {
        $manager = new Manager();
        $file = $this->getFile('test.csv');

        $reader = $manager->getReader($file, [
            'callback' => function ($reader) {
                $reader->setEnclosure('FOO');

                return $reader;
            }
        ]);
        $this->assertInstanceOf(PHPExcel_Reader_CSV::class, $reader);
        $this->assertEquals('FOO', $reader->getEnclosure());
    }

    public function testGetWriterXlsx()
    {
        $excel = $this->createMock(PHPExcel::class);
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $writer = $manager->getWriter($excel, $file);
        $this->assertInstanceOf(PHPExcel_Writer_Excel2007::class, $writer);
    }

    public function testGetWriterCsv()
    {
        $excel = $this->createMock(PHPExcel::class);
        $manager = new Manager();
        $file = $this->getFile('test.csv');

        $writer = $manager->getWriter($excel, $file, [
            'delimiter' => 'FOO'
        ]);
        $this->assertInstanceOf(PHPExcel_Writer_CSV::class, $writer);
        $this->assertEquals('FOO', $writer->getDelimiter());
    }

    public function testGetWriterCallback()
    {
        $excel = $this->createMock(PHPExcel::class);
        $manager = new Manager();
        $file = $this->getFile('test.csv');

        $writer = $manager->getWriter($excel, $file, [
            'callback' => function ($reader) {
                $reader->setEnclosure('FOO');
            }
        ]);
        $this->assertInstanceOf(PHPExcel_Writer_CSV::class, $writer);
        $this->assertEquals('FOO', $writer->getEnclosure());
    }

    public function testGetExcel()
    {
        $manager = new Manager();
        $file = $this->getFile('test.csv');

        $excel = $manager->getExcel($file);
        $this->assertInstanceOf(PHPExcel::class, $excel);
    }

    public function testRead()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $excel = $manager->getExcel($file);
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $results = $manager->read($worksheet, $table);

        $this->assertCount(5, $results);

        $first = $table->find()->first();

        $this->assertSame(1, $first->_row);
        $this->assertSame('a', $first->A);
        $this->assertSame('1', $first->B);
        $this->assertSame('1.01', $first->C);
        $this->assertSame('2017-01-01', $first->D);
        $this->assertSame('2017-01-01 01:00', $first->E);
        $this->assertSame('00:01:00', $first->F);
    }

    public function testReadLimitColumnsAndRows()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $excel = $manager->getExcel($file);
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $results = $manager->read($worksheet, $table, [
            'startColumn' => 'B',
            'endColumn' => 'D',
            'startRow' => 2,
            'endRow' => 3
        ]);

        $this->assertCount(2, $results);

        $first = $table->find()->first();

        $this->assertSame(1, $first->_row);
        $this->assertSame(null, $first->A);
        $this->assertSame('2', $first->B);
        $this->assertSame('2.02', $first->C);
        $this->assertSame('2017-01-02', $first->D);
        $this->assertSame(null, $first->E);
        $this->assertSame(null, $first->F);
    }

    public function testReadKeepOriginalRows()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $excel = $manager->getExcel($file);
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $manager->read($worksheet, $table, [
            'startRow' => 2,
            'keepOriginalRows' => true
        ]);

        $first = $table->find()->first();

        $this->assertSame(2, $first->_row);
    }

    public function testReadColumnMap()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $excel = $manager->getExcel($file);
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('MappedColumns');

        $results = $manager->read($worksheet, $table, [
            'columnMap' => [
                'A' => 'string_field',
                'B' => 'integer_field',
                'C' => 'float_field',
                'D' => 'date_field',
                'E' => 'datetime_field',
                'F' => 'time_field'
            ]
        ]);

        $this->assertCount(5, $results);

        $first = $table->find()->first();

        $this->assertSame(1, $first->id);
        $this->assertSame('a', $first->string_field);
        $this->assertSame(1, $first->integer_field);
        $this->assertSame(1.01, $first->float_field);
        $this->assertSame('2017-01-01', $first->date_field->format('Y-m-d'));
        $this->assertSame('2017-01-01 01:00', $first->datetime_field->format('Y-m-d H:i'));
        $this->assertSame('00:01:00', $first->time_field->format('H:i:s'));
    }

    public function testClear()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $excel = $manager->getExcel($file);
        $worksheet = $excel->getSheet();

        $manager->clear($worksheet);

        foreach ($worksheet->getRowIterator(1, 5) as $row) {
            foreach ($row->getCellIterator('A', 'F') as $cell) {
                $this->assertNull($cell->getValue());
            }
        }
    }

    public function testClearLimitRowsAndColumns()
    {
        $manager = new Manager();
        $file = $this->getFile('test.xlsx');

        $excel = $manager->getExcel($file);
        $worksheet = $excel->getSheet();

        $manager->clear($worksheet, [
            'startColumn' => 'B',
            'endColumn' => 'D',
            'startRow' => 2,
            'endRow' => 3
        ]);

        foreach ($worksheet->getRowIterator(4, 5) as $row) {
            foreach ($row->getCellIterator('A', 'F') as $cell) {
                $this->assertNotNull($cell->getValue());
            }
        }
        foreach ($worksheet->getRowIterator(2, 3) as $row) {
            foreach ($row->getCellIterator('E', 'F') as $cell) {
                $this->assertNotNull($cell->getValue());
            }
        }
        foreach ($worksheet->getRowIterator(2, 3) as $row) {
            foreach ($row->getCellIterator('B', 'D') as $cell) {
                $this->assertNull($cell->getValue());
            }
        }
    }

    public function testWrite()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $data = [
            1 => [
                '_row' => 1,
                'A' => 'a1',
                'B' => 'b1',
                'C' => 'c1',
                'D' => 1,
                'E' => 1.1,
                'F' => true
            ],
            2 => [
                '_row' => 10,
                'A' => 'a2',
                'B' => 'b2',
                'C' => 'c2',
                'D' => 2,
                'E' => 2.2,
                'F' => false
            ]
        ];

        $entities = $table->newEntities($data);
        $table->saveMany($entities);
        $this->assertCount(2, $table->find());

        $manager->write($table, $worksheet);

        foreach ($data as $row => $dataRow) {
            unset($dataRow['_row']);
            foreach ($dataRow as $column => $value) {
                $cell = $worksheet->getCell($column . $row);
                $this->assertEquals($value, $cell->getValue());
            }
        }
    }

    public function testWriteArray()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $data = [
            1 => [
                '_row' => 1,
                'A' => 'a1',
                'B' => 'b1',
                'C' => 'c1',
                'D' => 1,
                'E' => 1.1,
                'F' => true
            ],
            2 => [
                '_row' => 10,
                'A' => 'a2',
                'B' => 'b2',
                'C' => 'c2',
                'D' => 2,
                'E' => 2.2,
                'F' => false
            ]
        ];

        $entities = $table->newEntities($data);
        $table->saveMany($entities);
        $this->assertCount(2, $table->find());

        $table->eventManager()->on('Model.beforeFind', function ($e, $q) {
            return $q->hydrate(false);
        });

        $manager->write($table, $worksheet);

        foreach ($data as $row => $dataRow) {
            unset($dataRow['_row']);
            foreach ($dataRow as $column => $value) {
                $cell = $worksheet->getCell($column . $row);
                $this->assertEquals($value, $cell->getValue());
            }
        }
    }

    public function testWritePropertyMapAndCallbacks()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('MappedColumns');

        $data = [
            1 => [
                'id' => 1,
                'string_field' => 'A',
                'integer_field' => 1,
                'float_field' => 1.1,
                'date_field' => new Date('2017-01-01'),
                'datetime_field' => new Chronos('2017-01-01 00:01:00'),
                'time_field' => new Chronos('00:01:00'),
            ],
            2 => [
                'id' => 10,
                'string_field' => 'B',
                'integer_field' => 2,
                'float_field' => 1.2,
                'date_field' => new Date('2017-01-02'),
                'datetime_field' => new Chronos('2017-01-02 00:02:00'),
                'time_field' => new Chronos('00:02:00'),
            ]
        ];

        $entities = $table->newEntities($data);
        $table->saveMany($entities);
        $this->assertCount(2, $table->find());

        $map = [
            'string_field' => 'A',
            'integer_field' => 'B',
            'float_field' => 'C',
            'date_field' => 'D',
            'datetime_field' => 'E',
            'time_field' => 'F'
        ];

        $manager->write($table, $worksheet, [
            'propertyMap' => $map,
            'columnCallbacks' => [
                'D' => function ($cell) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('YYYY-MM-DD');
                },
                'E' => function ($cell) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('YYYY-MM-DD HH:MM:SS');
                },
                'F' => function ($cell) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('YYYY-MM-DD HH:MM:SS');
                }
            ]
        ]);

        foreach ($data as $row => $dataRow) {
            unset($dataRow['id']);
            foreach ($dataRow as $property => $value) {
                $column = $map[$property];
                $cell = $worksheet->getCell($column . $row);
                $this->assertEquals((string)$value, $cell->getFormattedValue());
            }
        }
    }

    public function testWriteKeepOriginalRows()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $data = [
            [
                '_row' => 1,
                'A' => 'a1',
                'B' => 'b1',
                'C' => 'c1',
                'D' => 'd1',
                'E' => 'e1',
                'F' => 'e1'
            ],
            [
                '_row' => 10,
                'A' => 'a2',
                'B' => 'b2',
                'C' => 'c2',
                'D' => 'd2',
                'E' => 'e2',
                'F' => 'e2'
            ],
            [
                '_row' => 100,
                'A' => 'a3',
                'B' => 'b3',
                'C' => 'c3',
                'D' => 'd3',
                'E' => 'e3',
                'F' => 'e3'
            ]
        ];

        $entities = $table->newEntities($data);
        $table->saveMany($entities);
        $this->assertCount(3, $table->find());

        $manager->write($table, $worksheet, [
            'keepOriginalRows' => true
        ]);

        foreach ($data as $dataRow) {
            $row = $dataRow['_row'];
            unset($dataRow['_row']);
            foreach ($dataRow as $column => $value) {
                $cell = $worksheet->getCell($column . $row);
                $this->assertEquals($value, $cell->getValue());
            }
        }
        $this->assertNull($worksheet->getCell('A2')->getValue());
    }

    /**
     * @expectedException UnexpectedValueException
     * @expectedExceptionMessage Cannot convert result to array.
     */
    public function testWriteWithFinderAndInvalidRecord()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $data = [
            1 => [
                '_row' => 1,
                'A' => 'a1',
                'B' => 'b1',
                'C' => 'c1',
                'D' => 'd1',
                'E' => 'e1',
                'F' => 'e1'
            ],
            2 => [
                '_row' => 10,
                'A' => 'a2',
                'B' => 'b2',
                'C' => 'c2',
                'D' => 'd2',
                'E' => 'e2',
                'F' => 'e2'
            ]
        ];

        $entities = $table->newEntities($data);
        $table->saveMany($entities);
        $this->assertCount(2, $table->find());

        $manager->write($table, $worksheet, [
            'finder' => 'list'
        ]);
    }

    public function testSaveAndCallbackWriter()
    {
        $writer = $this->createMock(PHPExcel_Writer_IWriter::class);
        $excel = $this->createMock(PHPExcel::class);
        $file = $this->getFile('test.xlsx');

        $writer->expects($this->once())
            ->method('save')
            ->with($file->pwd());

        $manager = new Manager();

        $manager->save($excel, $file, [
            'callback' => function () use ($writer) {
                return $writer;
            }
        ]);
    }

    public function testWriteAndAttachHeader()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $manager->write($table, $worksheet, [
            'startRow' => 2,
            'header' => [
                'a' => 'Foo',
                'b' => 'Bar'
            ]
        ]);

        $cellA = $worksheet->getCell('A1');
        $this->assertEquals('Foo', $cellA->getValue());
        $this->assertTrue($cellA->getStyle()->getFont()->getBold());

        $cellB = $worksheet->getCell('B1');
        $this->assertEquals('Bar', $cellB->getValue());
        $this->assertTrue($cellB->getStyle()->getFont()->getBold());
    }

    /**
     * @expectedException LogicException
     * @expectedExceptionMessage Option `startRow` must be > 1 if you want to attach header.
     */
    public function testWriteAndAttachHeaderInvalidStartRow()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();
        $table = TableRegistry::get('RegularColumns');

        $manager->write($table, $worksheet, [
            'header' => [
                'a' => 'Foo',
                'b' => 'Bar'
            ]
        ]);
    }

    public function testAttachHeader()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();

        $header = [
            'a' => 'Foo',
            'b' => 'Bar'
        ];

        $result = $manager->attachHeader($worksheet, $header);
        $this->assertSame($worksheet, $result);

        $cellA = $worksheet->getCell('A1');
        $this->assertEquals('Foo', $cellA->getValue());
        $this->assertTrue($cellA->getStyle()->getFont()->getBold());

        $cellB = $worksheet->getCell('B1');
        $this->assertEquals('Bar', $cellB->getValue());
        $this->assertTrue($cellB->getStyle()->getFont()->getBold());
    }

    public function testAttachHeaderCustomOptions()
    {
        $manager = new Manager();
        $excel = new PHPExcel();
        $worksheet = $excel->getSheet();

        $header = [
            'a' => 'Foo',
            'b' => 'Bar'
        ];

        $result = $manager->attachHeader($worksheet, $header, [
            'row' => 2,
            'style' => [
                'font' => [
                    'size' => 100
                ]
            ]
        ]);
        $this->assertSame($worksheet, $result);

        $cellA = $worksheet->getCell('A2');
        $this->assertEquals('Foo', $cellA->getValue());
        $this->assertEquals(100, $cellA->getStyle()->getFont()->getSize());

        $cellB = $worksheet->getCell('B2');
        $this->assertEquals('Bar', $cellB->getValue());
        $this->assertEquals(100, $cellB->getStyle()->getFont()->getSize());
    }
}
