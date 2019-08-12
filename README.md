# Spout

## Fork Notes

This is a somewhat messy fork of the Box/Spout library. It is forked from a version of the Spout 3.0 development branch, though we'll eventually try and catch back up to their recent master release.

At a high-level, we try to implement features we desired in a way that was largely compatible with upstream code, and that maintained Spout's general pattern of being light on memory and reasonably efficient. As we mostly focus on XLSX export functionality, no changes have been made to ODS or CSV functionality, nor have any changes been made with how XLSX files are read.

With this in mind, the following is the list of notable changes in this fork:

  1. Worksheet adds a `getOption` and `setOption` method, which is used to set various arbitrary options on worksheets.
  2. Cell style merging logic has been optimized. It was slow.
  3. A cell type has been added for images. When used, this will embed an image into the cell without overflowing into other cells. The cell's value should be set to the URL of an image. We use Guzzle with batch downloading to try to speed up this functionality as much as possible.
  4. Frozen panes have been added. The `freeze_pane` worksheet option can be passed an x and y coordinate for where to freeze. For instance, `$worksheet->setOption('freeze_pane', [2, 2])` would freeze at the C3 cell.
  5. Worksheet column widths can be set with the `column_widths` option. This should be an array of widths, one for each column. For instance, `$worksheet->setOption('column_widths', [10, 20, 40])`. The value `'auto'` can be used to perform automatic column width calculation, as in `$worksheet->setOption('column_widths', [10, 20, 'auto'])`; this is implemented using a very lazy algorithm for speed -- we simply find the greatest number of characters in any given cell to determine what width should be used for a column.
  6. Autofiltering can be enabled for all columns by setting the `filter` option, e.g. `$worksheet->setOption('filter', true)`.


## Documentation

Full documentation can be found at [http://opensource.box.com/spout/](http://opensource.box.com/spout/).


## Requirements

* PHP version 5.6 or higher
* PHP extension `php_zip` enabled
* PHP extension `php_xmlreader` enabled


## Running tests

On the `master` branch, only unit and functional tests are included. The performance tests require very large files and have been excluded.
If you just want to check that everything is working as expected, executing the tests of the `master` branch is enough.

If you want to run performance tests, you will need to checkout the `perf-tests` branch. Multiple test suites can then be run, depending on the expected output:

* `phpunit` - runs the whole test suite (unit + functional + performance tests)
* `phpunit --exclude-group perf-tests` - only runs the unit and functional tests
* `phpunit --group perf-tests` - only runs the performance tests

For information, the performance tests take about 30 minutes to run (processing 1 million rows files is not a quick thing).

> Performance tests status: [![Build Status](https://travis-ci.org/box/spout.svg?branch=perf-tests)](https://travis-ci.org/box/spout)


## Support

You can ask questions, submit new features ideas or discuss about Spout in the chat room:<br>
[![Gitter](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/box/spout?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

## Copyright and License

Copyright 2017 Box, Inc. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

   http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
