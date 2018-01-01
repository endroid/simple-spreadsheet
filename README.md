# Simple Excel

*By [endroid](https://endroid.nl/)*

[![Latest Stable Version](http://img.shields.io/packagist/v/endroid/simple-excel.svg)](https://packagist.org/packages/endroid/simple-excel)
[![Build Status](http://img.shields.io/travis/endroid/simple-excel.svg)](http://travis-ci.org/endroid/simple-excel)
[![Total Downloads](http://img.shields.io/packagist/dt/endroid/simple-excel.svg)](https://packagist.org/packages/endroid/simple-excel)
[![License](http://img.shields.io/packagist/l/endroid/simple-excel.svg)](https://packagist.org/packages/endroid/simple-excel)

Library for quickly loading and generating Excel files. Data can be loaded
from and converted to an array, an Excel/CSV file or PHPExcel object.

Great advantage of this library is the small amount of code needed to perform
an import or export of data, given one of the above formats.

## Installation

Use [Composer](https://getcomposer.org/) to install the library.

``` bash
$ composer require endroid/simple-excel
```

## Usage

```php
<?php

use Endroid\SimpleExcel\SimpleExcel;

$excel = new SimpleExcel();
$excel->loadFromFile(__DIR__.'/../Resources/data/data.xlsx');
$excel->loadFromArray([
    'Players' => [
        ['name' => 'L. Messi', 'club' => 'Barcelona'],
        ['name' => 'C. Ronaldo', 'club' => 'Real Madrid']
    ]
]);

$excel->saveToOutput('players.csv', ['Players']);
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
