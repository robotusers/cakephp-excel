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
namespace Robotusers\Excel\Test\TestCase\Model\Table;

use Cake\ORM\TableRegistry;
use Robotusers\Excel\Model\Behavior\ExcelBehavior;
use Robotusers\Excel\Model\Sheet;
use Robotusers\Excel\Test\TestCase;

/**
 * Description of SheetTest
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
class SheetTest extends TestCase
{
    public function testInitialize()
    {
        $config = [
            'foo' => 'bar'
        ];

        $table = TableRegistry::get('MappedColumns', [
            'className' => Sheet::class,
            'excel' => $config
        ]);

        $behavior = $table->behaviors()->get('Excel');
        $this->assertInstanceOf(ExcelBehavior::class, $behavior);
        $this->assertEquals('bar', $behavior->getConfig('foo'));
    }
}