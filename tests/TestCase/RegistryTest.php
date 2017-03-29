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
namespace Robotusers\Excel\Test\TestCase;

use Cake\Database\Connection;
use Cake\Database\Driver\Sqlite;
use Cake\Database\Schema\TableSchema;
use Cake\Datasource\ConnectionManager;
use Cake\ORM\Locator\LocatorInterface;
use PHPExcel;
use PHPExcel_Reader_IReader;
use PHPExcel_Worksheet;
use Robotusers\Excel\Database\Factory;
use Robotusers\Excel\Excel\Manager;
use Robotusers\Excel\Model\Sheet;
use Robotusers\Excel\Registry;
use Robotusers\Excel\Test\TestCase;

/**
 * Description of RegistryTest
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class RegistryTest extends TestCase
{

    /**
     *
     * @return Registry
     */
    protected function createRegistry()
    {
        $manager = $this->createMock(Manager::class);
        $factory = $this->createMock(Factory::class);
        $connection = $this->createMock(Connection::class);
        $locator = $this->createMock(LocatorInterface::class);

        $registry = new Registry($manager, $factory);
        $registry->setConnection($connection);
        $registry->tableLocator($locator);

        return $registry;
    }

    public function testGet()
    {
        $file = $this->getFile('test.xlsx');
        $registry = $this->createRegistry();

        $name = 'Results';
        $locatorOptions = [
            'table' => 'foo'
        ];
        $options = [
            'foo' => 'bar'
        ];

        $reader = $this->createMock(PHPExcel_Reader_IReader::class);
        $excel = $this->createMock(PHPExcel::class);
        $worksheet = $this->createMock(PHPExcel_Worksheet::class);
        $schema = $this->createMock(TableSchema::class);
        $table = $this->getMockBuilder(Sheet::class)
            ->disableOriginalConstructor()
            ->setMethods([
                'setSchema',
                'setFile',
                'setWorksheet',
                'readExcel'
            ])
            ->getMock();

        $registry->getManager()
            ->expects($this->once())
            ->method('getReader')
            ->with($file)
            ->willReturn($reader);

        $reader->expects($this->once())
            ->method('load')
            ->with($file->pwd())
            ->willReturn($excel);

        $excel->method('sheetNameExists')
            ->with('foo')
            ->willReturn(true);

        $excel->method('getSheetByName')
            ->with('foo')
            ->willReturn($worksheet);

        $registry->getFactory()
            ->expects($this->at(0))
            ->method('createSchema')
            ->with($worksheet)
            ->willReturn($schema);

        $schema->method('name')
            ->willReturn($name);

        $registry->getFactory()
            ->expects($this->at(1))
            ->method('createTable')
            ->with($registry->getConnection(), $schema);

        $registry->tableLocator()
            ->method('get')
            ->with($name, [
                'className' => Sheet::class,
                'connection' => $registry->getConnection(),
                'excel' => $options,
                'table' => 'foo'
            ])
            ->willReturn($table);

        $table->expects($this->any())
            ->method('setSchema')
            ->with($schema)
            ->willReturn($table);
        $table->expects($this->any())
            ->method('setFile')
            ->with($file)
            ->willReturn($table);
        $table->expects($this->any())
            ->method('setWorksheet')
            ->with($worksheet)
            ->willReturn($table);
        $table->expects($this->any())
            ->method('readExcel');

        $sheet = $registry->get($file, 'foo', $options, $locatorOptions);

        $this->assertSame($table, $sheet);
    }

    /**
     * @covers \Robotusers\Excel\Registry::instance
     * @covers \Robotusers\Excel\Registry::getManager
     * @covers \Robotusers\Excel\Registry::getFactory
     * @covers \Robotusers\Excel\Registry::getConnection
     */
    public function testDefaultInstance()
    {
        ConnectionManager::dropAlias('excel');
        $registry = Registry::instance();
        $this->assertInstanceOf(Registry::class, $registry);
        $this->assertInstanceOf(Manager::class, $registry->getManager());
        $this->assertInstanceOf(Factory::class, $registry->getFactory());
        $this->assertInstanceOf(Connection::class, $registry->getConnection());
        $this->assertInstanceOf(Sqlite::class, $registry->getConnection()->getDriver());
    }
}
