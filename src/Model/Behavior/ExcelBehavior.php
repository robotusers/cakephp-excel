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

namespace Robotusers\Excel\Model\Behavior;

use ArrayObject;
use Cake\Datasource\EntityInterface;
use Cake\Event\Event;
use Cake\Filesystem\File;
use Cake\ORM\Behavior;
use Cake\ORM\Table;
use InvalidArgumentException;
use PHPExcel;
use PHPExcel_Worksheet;
use Robotusers\Excel\Excel\Manager;
use Robotusers\Excel\Traits\DiscoverWorksheetTrait;
use RuntimeException;

/**
 * Description of ExcelBehavior
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class ExcelBehavior extends Behavior
{

    use DiscoverWorksheetTrait;

    /**
     *
     * @var File
     */
    protected $file;

    /**
     *
     * @var PHPExcel_Worksheet
     */
    protected $worksheet;

    /**
     *
     * @var Manager
     */
    protected $manager;

    /**
     *
     * @var array
     */
    protected $_defaultConfig = [
        'startRow' => 1,
        'endRow' => null,
        'startColumn' => 'A',
        'endColumn' => null,
        'columnMap' => [],
        'propertyMap' => [],
        'finder' => 'all',
        'finderOptions' => [],
        'marshallerOptions' => [],
        'saveOptions' => []
    ];

    /**
     *
     * @param array $config
     * @return void
     */
    public function initialize(array $config)
    {
        if (isset($config['manager'])) {
            if (!$config['manager'] instanceof Manager) {
                throw new InvalidArgumentException('Invalid manager.');
            }
            $this->manager = $config['manager'];
        }

        if (!isset($config['propertyMap'])) {
            $propertyMap = array_flip($this->getConfig('columnMap'));
            $this->setConfig('propertyMap', $propertyMap);
        }
    }

    /**
     *
     * @param array $options
     * @return EntityInterface[]
     */
    public function readExcel(array $options = [])
    {
        $options += $this->getConfig();
        $worksheet = $this->getWorksheet();

        return $this->getManager()->read($worksheet, $this->_table, $options);
    }

    /**
     *
     * @param array $options
     * @return File
     */
    public function writeExcel(array $options = [])
    {
        $options += $this->getConfig();

        $worksheet = $this->getWorksheet();
        $manager = $this->getManager();
        $manager->clear($worksheet, $options);
        $manager->write($this->_table, $worksheet, $options);

        $file = $this->getFile();
        $writer = $manager->getWriter($worksheet->getParent(), $file);

        $writer->save($file->pwd());

        return $file;
    }

    /**
     *
     * @return Manager
     */
    public function getManager()
    {
        if ($this->manager === null) {
            $this->manager = new Manager();
        }

        return $this->manager;
    }

    /**
     *
     * @return PHPExcel_Worksheet
     * @throws RuntimeException
     */
    public function getWorksheet()
    {
        if ($this->worksheet === null) {
            throw new RuntimeException('Worksheet has not been set.');
        }

        return $this->worksheet;
    }

    /**
     *
     * @param string|int|PHPExcel_Worksheet $worksheet
     * @param array $options
     * @return Table
     */
    public function setWorksheet($worksheet = null, array $options = [])
    {
        if (!$worksheet instanceof PHPExcel_Worksheet) {
            $file = $this->getFile();
            $excel = $this->getManager()->getExcel($file, $options);
            $worksheet = $this->discoverWorksheet($excel, $worksheet);
        }
        $this->worksheet = $worksheet;

        return $this->_table;
    }

    /**
     *
     * @return PHPExcel
     */
    public function getExcel()
    {
        return $this->getWorksheet()->getParent();
    }

    /**
     *
     * @return File
     * @throws RuntimeException
     */
    public function getFile()
    {
        if ($this->file === null) {
            throw new RuntimeException('File has not been set.');
        }

        return $this->file;
    }

    /**
     *
     * @param File $file
     * @return Table
     */
    public function setFile(File $file)
    {
        $this->file = $file;

        return $this->_table;
    }
}
