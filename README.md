# Simple Spreadsheet

*By [endroid](https://endroid.nl/)*

[![Latest Stable Version](http://img.shields.io/packagist/v/endroid/simple-spreadsheet.svg)](https://packagist.org/packages/endroid/simple-spreadsheet)
[![Build Status](https://github.com/endroid/simple-spreadsheet/workflows/CI/badge.svg)](https://github.com/endroid/simple-spreadsheet/actions)
[![Total Downloads](http://img.shields.io/packagist/dt/endroid/simple-spreadsheet.svg)](https://packagist.org/packages/endroid/simple-spreadsheet)
[![License](http://img.shields.io/packagist/l/endroid/simple-spreadsheet.svg)](https://packagist.org/packages/endroid/simple-spreadsheet)

Library for quickly importing and exporting spreadsheet data. Data can be loaded
from and converted from / to an array, an Excel/CSV file or Spreadsheet object.

The main advantage of this library is the small amount of code needed to perform
an import or export of data, given one of the above formats.

## Installation

Use [Composer](https://getcomposer.org/) to install the library.

``` bash
$ composer require endroid/simple-spreadsheet
```

## Usage

```php
<?php

use Endroid\SimpleSpreadsheet\Adapter\FileAdapter;
use Endroid\SimpleSpreadsheet\SimpleSpreadsheet;

$spreadsheet = new SimpleSpreadsheet();
$spreadsheet->load('data.xlsx'); // Load all sheets from data.xlsx
$spreadsheet->load([
    'Players' => [
        ['name' => 'L. Messi', 'club' => 'Barcelona'],
        ['name' => 'C. Ronaldo', 'club' => 'Real Madrid']
    ]
]);

$spreadsheet->save(FileAdapter::class, ['Players'], ['filename' => 'players.csv']);
```

You can also use the saveToString and getHeadersByFilename methods to build a
Response object instead of directly outputting to the browser.

## Versioning

Version numbers follow the MAJOR.MINOR.PATCH scheme. Backwards compatible
changes will be kept to a minimum but be aware that these can occur. Lock
your dependencies for production and test your code when upgrading.

## License

This bundle is under the MIT license. For the full copyright and license
information please view the LICENSE file that was distributed with this source code.
