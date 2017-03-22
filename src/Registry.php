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

namespace Robotusers\Excel;

use Cake\Database\Connection;
use Cake\Datasource\ConnectionInterface;
use Cake\Datasource\ConnectionManager;
use Cake\Filesystem\File;
use Cake\ORM\Locator\LocatorAwareTrait;
use Cake\Utility\Inflector;
use PHPExcel_Worksheet;
use Robotusers\Excel\Database\Factory;
use Robotusers\Excel\Excel\Manager;
use Robotusers\Excel\Model\Sheet;
use Robotusers\Excel\Traits\DiscoverWorksheetTrait;

/**
 * Description of SheetLoader
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class Registry
{
    const CONNECTON_NAME = 'excel';

    use DiscoverWorksheetTrait;
    use LocatorAwareTrait;

    /**
     *
     * @var Manager
     */
    protected $manager;

    /**
     *
     * @var Connection
     */
    protected $connection;

    /**
     *
     * @var Factory
     */
    protected $factory;

    /**
     *
     * @var array
     */
    protected $sheets = [];

    /**
     *
     * @var self
     */
    protected static $instance;

    /**
     *
     * @param Manager $manager
     * @param Factory $factory
     */
    public function __construct(Manager $manager, Factory $factory)
    {
        $this->manager = $manager;
        $this->factory = $factory;
    }

    /**
     *
     * @param string|File $file
     * @param string $sheet
     * @param array $options
     * @param array $locatorOptions
     * @return Sheet
     */
    public function get($file, $sheet = null, array $options = [], array $locatorOptions = [])
    {
        if (!$file instanceof File) {
            $file = new File($file);
        }
        if (is_array($sheet)) {
            $options = $sheet;
            $sheet = null;
        }

        $reader = $this->manager->getReader($file, $options);
        $excel = $reader->load($file->pwd());
        $worksheet = $this->discoverWorksheet($excel, $sheet);

        $hash = $file->md5();
        $sheetId = $excel->getIndex($worksheet);

        if (!isset($this->sheets[$hash][$sheetId])) {
            $locatorOptions += [
                'excel' => $options
            ];
            $table = $this->loadSheet($file, $worksheet, $locatorOptions);

            $this->sheets[$hash][$sheetId] = $table;
        }

        return $this->sheets[$hash][$sheetId];
    }

    /**
     *
     * @param File $file
     * @param PHPExcel_Worksheet $worksheet
     * @param array $options
     * @return Sheet
     */
    protected function loadSheet(File $file, PHPExcel_Worksheet $worksheet, array $options)
    {
        $schema = $this->factory->createSchema($worksheet, $options['excel']);
        $connection = $this->getConnection();
        $this->factory->createTable($connection, $schema);

        $name = $schema->name();
        $alias = Inflector::camelize($name);

        $options += [
            'className' => Sheet::class,
            'connection' => $connection
        ];

        $table = $this->tableLocator()->get($alias, $options);
        $table->setSchema($schema)
            ->setFile($file)
            ->setWorksheet($worksheet)
            ->readExcel();

        return $table;
    }

    /**
     *
     * @return ConnectionInterface
     */
    public function getConnection()
    {
        if ($this->connection === null) {
            $this->connection = ConnectionManager::get(static::CONNECTON_NAME);
        }

        return $this->connection;
    }

    /**
     *
     * @param Connection $connection
     * @return $this
     */
    public function setConnection(Connection $connection)
    {
        $this->connection = $connection;

        return $this;
    }

    /**
     *
     * @return Factory
     */
    public function getFactory()
    {
        return $this->factory;
    }

    /**
     *
     * @return Manager
     */
    public function getManager()
    {
        return $this->manager;
    }

    /**
     *
     * @return self
     */
    public static function instance()
    {
        if (static::$instance === null) {
            $manager = new Manager();
            $factory = new Factory();
            static::$instance = new self($manager, $factory);
        }

        return static::$instance;
    }
}
