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
namespace Robotusers\Excel\Test\TestCase\Model\Behavior;

use Cake\Datasource\EntityInterface;
use Cake\Event\EventList;
use Cake\Event\EventManager;
use Cake\Filesystem\File;
use Cake\ORM\Table;
use Cake\ORM\TableRegistry;
use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Robotusers\Excel\Excel\Manager;
use Robotusers\Excel\Test\TestCase;
use RuntimeException;
use stdClass;

/**
 * Description of ExcelBehaviorTest
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class ExcelBehaviorTest extends TestCase
{
    public $fixtures = [
        'plugin.Robotusers/Excel.mapped_columns'
    ];

    public function setUp()
    {
        parent::setUp();
        EventManager::instance()->setEventList(new EventList);
    }

    /**
     *
     * @param array $options
     * @return Table
     */
    protected function createTable(array $options = [])
    {
        $table = TableRegistry::get('MappedColumns');
        $table->addBehavior('Robotusers/Excel.Excel', $options + [
            'columnMap' => [
                'A' => 'string_field',
                'B' => 'integer_field',
                'C' => 'float_field',
                'D' => 'date_field',
                'E' => 'datetime_field',
                'F' => 'time_field'
            ]
        ]);

        return $table;
    }

    public function testPropertyMap()
    {
        $table = TableRegistry::get('MappedColumns');
        $table->addBehavior('Robotusers/Excel.Excel', [
            'columnMap' => [
                '*' => false,
                'A' => true,
                'B' => false,
                'C' => 'float_field',
                'D' => 'date_field',
                'E' => 'datetime_field',
                'F' => 'time_field'
            ]
        ]);

        $expected = [
            '*' => false,
            'A' => true,
            'B' => false,
            'float_field' => 'C',
            'date_field' => 'D',
            'datetime_field' => 'E',
            'time_field' => 'F'
        ];

        $map = $table->behaviors()->Excel->getConfig('propertyMap');
        $this->assertEquals($expected, $map);
    }

    /**
     * @covers \Robotusers\Excel\Model\Behavior\ExcelBehavior::getFile
     * @covers \Robotusers\Excel\Model\Behavior\ExcelBehavior::setFile
     */
    public function testFile()
    {
        $table = $this->createTable();
        $file = $this->getFile('test.xlsx');

        $table->setFile($file);
        $tableFile = $table->getFile();

        $this->assertSame($file, $tableFile);
    }

    /**
     * @expectedException RuntimeException
     * @expectedExceptionMessage File has not been set.
     */
    public function testFileMissing()
    {
        $table = $this->createTable();
        $table->getFile();
    }

    /**
     * @covers \Robotusers\Excel\Model\Behavior\ExcelBehavior::setWorksheet
     * @covers \Robotusers\Excel\Model\Behavior\ExcelBehavior::getWorksheet
     */
    public function testWorksheet()
    {
        $table = $this->createTable();
        $worksheet = $this->createMock(Worksheet::class);

        $table->setWorksheet($worksheet);
        $tableWorksheet = $table->getWorksheet();

        $this->assertSame($worksheet, $tableWorksheet);
    }

    public function testWorksheetName()
    {
        $excel = $this->createMock(Spreadsheet::class);
        $manager = $this->createMock(Manager::class);
        $worksheet = $this->createMock(Worksheet::class);
        $file = $this->createMock(File::class);

        $table = $this->createTable([
            'manager' => $manager
        ]);
        $table->setFile($file);

        $manager->method('getExcel')
            ->willReturn($excel);

        $excel->expects($this->once())
            ->method('sheetNameExists')
            ->with('foo')
            ->willReturn(true);

        $excel->expects($this->once())
            ->method('getSheetByName')
            ->with('foo')
            ->willReturn($worksheet);

        $table->setWorksheet('foo');
        $tableWorksheet = $table->getWorksheet();

        $this->assertSame($worksheet, $tableWorksheet);
    }

    public function testWorksheetCodeName()
    {
        $excel = $this->createMock(Spreadsheet::class);
        $manager = $this->createMock(Manager::class);
        $worksheet = $this->createMock(Worksheet::class);
        $file = $this->createMock(File::class);

        $table = $this->createTable([
            'manager' => $manager
        ]);
        $table->setFile($file);

        $manager->method('getExcel')
            ->willReturn($excel);

        $excel->expects($this->once())
            ->method('sheetNameExists')
            ->with('foo')
            ->willReturn(false);

        $excel->expects($this->once())
            ->method('sheetCodeNameExists')
            ->with('foo')
            ->willReturn(true);

        $excel->expects($this->once())
            ->method('getSheetByCodeName')
            ->with('foo')
            ->willReturn($worksheet);

        $table->setWorksheet('foo');
        $tableWorksheet = $table->getWorksheet();

        $this->assertSame($worksheet, $tableWorksheet);
    }

    public function testWorksheetIndex()
    {
        $excel = $this->createMock(Spreadsheet::class);
        $manager = $this->createMock(Manager::class);
        $worksheet = $this->createMock(Worksheet::class);
        $file = $this->createMock(File::class);

        $table = $this->createTable([
            'manager' => $manager
        ]);
        $table->setFile($file);

        $manager->method('getExcel')
            ->willReturn($excel);

        $excel->expects($this->once())
            ->method('sheetNameExists')
            ->with(1)
            ->willReturn(false);

        $excel->expects($this->once())
            ->method('sheetCodeNameExists')
            ->with(1)
            ->willReturn(false);

        $excel->expects($this->once())
            ->method('getSheet')
            ->with(1)
            ->willReturn($worksheet);

        $table->setWorksheet(1);
        $tableWorksheet = $table->getWorksheet();

        $this->assertSame($worksheet, $tableWorksheet);
    }

    public function testWorksheetActive()
    {
        $excel = $this->createMock(Spreadsheet::class);
        $manager = $this->createMock(Manager::class);
        $worksheet = $this->createMock(Worksheet::class);
        $file = $this->createMock(File::class);

        $table = $this->createTable([
            'manager' => $manager
        ]);
        $table->setFile($file);

        $manager->method('getExcel')
            ->willReturn($excel);

        $excel->expects($this->never())
            ->method('sheetNameExists');

        $excel->expects($this->never())
            ->method('sheetCodeNameExists');

        $excel->expects($this->once())
            ->method('getActiveSheet')
            ->willReturn($worksheet);

        $table->setWorksheet(null);
        $tableWorksheet = $table->getWorksheet();

        $this->assertSame($worksheet, $tableWorksheet);
    }

    /**
     * @expectedException RuntimeException
     * @expectedExceptionMessage Worksheet has not been set.
     */
    public function testWorksheetMissing()
    {
        $table = $this->createTable();
        $table->getWorksheet();
    }

    /**
     * @covers \Robotusers\Excel\Model\Behavior\ExcelBehavior::getExcel
     */
    public function testGetExcel()
    {
        $table = $this->createTable();
        $excel = $this->createMock(Spreadsheet::class);
        $worksheet = $this->createMock(Worksheet::class);
        $worksheet->method('getParent')->willReturn($excel);

        $table->setWorksheet($worksheet);
        $tableExcel = $table->getExcel();

        $this->assertSame($excel, $tableExcel);
    }

    /**
     * @covers \Robotusers\Excel\Model\Behavior\ExcelBehavior::getManager
     */
    public function testGetManager()
    {
        $manager = $this->createMock(Manager::class);
        $table = $this->createTable([
            'manager' => $manager
        ]);

        $tableManager = $table->getManager();

        $this->assertSame($manager, $tableManager);
    }

    /**
     * @expectedException InvalidArgumentException
     * @expectedExceptionMessage Invalid manager.
     */
    public function testInvalidManager()
    {
        $manager = new stdClass();
        $table = $this->createTable([
            'manager' => $manager
        ]);

        $table->getManager();
    }

    public function testReadExcel()
    {
        $manager = $this->createMock(Manager::class);
        $worksheet = $this->createMock(Worksheet::class);
        $table = $this->createTable([
            'manager' => $manager
        ]);
        $options = [
            'foo' => 'bar'
        ];
        $results = [
            $this->createMock(EntityInterface::class)
        ];

        $manager->expects($this->once())
            ->method('read')
            ->with($worksheet, $table, $options + $table->behaviors()->get('Excel')->getConfig())
            ->willReturn($results);

        $table->setWorksheet($worksheet);
        $read = $table->readExcel($options);
        $this->assertEquals($results, $read);

        $this->assertEventFired('Model.beforeReadWorksheet');
    }

    public function testWriteExcel()
    {
        $file = $this->getFile('test_empty.xlsx', true);
        $table = $this->createTable();
        $table->setFile($file);

        $excel = $table->getManager()->getExcel($file);
        $table->setWorksheet($excel->getActiveSheet());

        $data = [
            1 => [
                'string_field' => 'A',
                'integer_field' => 1,
                'float_field' => 1.1,
                'date_field' => '2017-01-01',
                'datetime_field' => '2017-01-01 00:01:00',
                'time_field' => '00:01:00',
            ],
            2 => [
                'string_field' => 'B',
                'integer_field' => 2,
                'float_field' => 1.2,
                'date_field' => '2017-01-02',
                'datetime_field' => '2017-01-02 00:02:00',
                'time_field' => '00:02:00',
            ]
        ];
        $entities = $table->newEntities($data);
        $table->saveMany($entities);

        $table->writeExcel([
            'columnCallbacks' => [
                'D' => function ($cell) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('YYYY-MM-DD');
                },
                'E' => function ($cell) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('YYYY-MM-DD HH:MM:SS');
                },
                'F' => function ($cell) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('HH:MM:SS');
                }
            ]
        ]);

        $writtenExcel = $table->getManager()->getExcel($file);
        $map = $table->behaviors()->get('Excel')->getConfig('columnMap');

        $worksheet = $writtenExcel->getActiveSheet();
        foreach ($worksheet->getRowIterator(1, 2) as $i => $row) {
            foreach ($row->getCellIterator('A', 'F') as $cell) {
                $property = $map[$cell->getColumn()];

                $this->assertEquals($data[$i][$property], $cell->getFormattedValue());
            }
        }

        $file->delete();

        $this->assertEventFired('Model.beforeWriteWorksheet');
        $this->assertEventFired('Model.beforeWriteExcel');
    }
}
