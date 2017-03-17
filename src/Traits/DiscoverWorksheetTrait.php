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

namespace Robotusers\Excel\Traits;

use InvalidArgumentException;
use PHPExcel;
use PHPExcel_Worksheet;

/**
 *
 * @author Robert Pustułka <r.pustulka@robotusers.com>
 */
trait DiscoverWorksheetTrait
{

    /**
     *
     * @param PHPExcel $excel
     * @param string $name
     * @return PHPExcel_Worksheet
     * @throws InvalidArgumentException
     */
    protected function discoverWorksheet(PHPExcel $excel, $name = null)
    {
        if ($name !== null) {
            if ($excel->sheetNameExists($name)) {
                $worksheet = $excel->getSheetByName($name);
            } elseif ($excel->sheetCodeNameExists($name)) {
                $worksheet = $excel->getSheetByCodeName($name);
            } elseif (is_int($name)) {
                $worksheet = $excel->getSheet($name);
            } else {
                throw new InvalidArgumentException('Sheet does not exist.');
            }
        } else {
            $worksheet = $excel->getActiveSheet();
        }

        return $worksheet;
    }
}