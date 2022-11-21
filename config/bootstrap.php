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

use Cake\Database\Connection;
use Cake\Database\Driver\Sqlite;
use Cake\Datasource\ConnectionManager;
use Cake\Log\Log;
use Robotusers\Excel\Registry;

$name = Registry::CONNECTON_NAME;
$hasConnectionConfig = ConnectionManager::getConfig($name);
if (!$hasConnectionConfig && !in_array('sqlite', PDO::getAvailableDrivers())) {
    $msg = 'Spreadsheet not enabled. You need to either install pdo_sqlite, or define the "%s" connection name.';
    Log::warning(sprintf($msg, $name));
    return;
}
if (!$hasConnectionConfig) {
    ConnectionManager::setConfig($name, [
        'className' => Connection::class,
        'driver' => Sqlite::class,
        'database' => ':memory:',
        'encoding' => 'utf8',
        'cacheMetadata' => true,
        'quoteIdentifiers' => true,
    ]);
}
