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
namespace Robotusers\Excel\Test;

use Cake\TestSuite\TestCase as CakeTestCase;
use InvalidArgumentException;
use SplFileInfo;

/**
 * Description of TestCase
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class TestCase extends CakeTestCase
{

    /**
     *
     * @param string $name
     * @param bool $create
     * @return File
     * @throws InvalidArgumentException
     */
    protected function getFile($name, $create = false)
    {
        $path = PLUGIN_ROOT . DS . 'tests' . DS . 'files' . DS . $name;

        $file = new File($path, $create);
        if (!$file->exists()) {
            $message = sprintf('Missing file "%s".', $name);
            throw new InvalidArgumentException($message);
        }

        return $file;
    }

    /**
     *
     * @param string $filename
     * @param int $sheetIndex
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    protected function loadWorksheet($filename, $sheetIndex = 0)
    {
        $file = $this->getFile($filename);

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file->getRealPath());
        $excel = $reader->load($file->getRealPath());

        return $excel->getSheet($sheetIndex);
    }
}
