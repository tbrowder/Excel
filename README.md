[![Build Status](https://travis-ci.com/tbrowder/Excel-Raku.svg?branch=master)](https://travis-ci.com/tbrowder/Excel-Raku)

# Excel

A Raku module to create or use Excel xlsx files. It
can:

* Read an existing Microsoft Excel xlsx file
* Create Excel xlxs files
* Use an Excel file as a template to create new Excel xlsx files

Planned:

* Use a text file as a template to create new Excel xlsx files

## DESCRIPTION

This module provides the capability of using Excel and text templates
to generate Excel files. It is a WIP and has little working code
at the moment. If you are interested in the concept, please
star the project, follow it, and file a feature request issue.

The project will use several Perl modules which will have to
be installed for the distro to work (I use `cpanm` for that):

+ Perl modules required:

    + `Excel::Writer::XLSX`
    + `Spreadsheet::XLSX`

## GENESIS

This project started when I was trying to automate creating forms for
my tax return. I have a need to generate multiple workbooks, one
worksheet per workbook, from a template, and am designing my own
format for that purpose. I will use the `Raku` language to parse the
text-file template, then, with the aid of the `Raku` module
`Inline::Perl5`, I will read my xlsx data files with one of the Perl
xlsx readers and then use this module to write new, filtered files in
the form of the template.

I am just starting, but I'm looking at a template format something like
this, one line per row, cells separated by pipes (`|`), key/value
attribute pairs (using a syntax similar to `Raku`s Pairs) following the cell content:

``` Raku
# This is a comment. The following row describes one xlsx row with four columns (the
# first column being empty) and it has an ending comment.
# Comments are stripped to the end-of-line eol before parsing the row.

| some text | 5.26 | :formula<some formula> :color<red> :width(2) # comment...

# Empty rows are ignored, the worksheet will have all rows padded with empty cells
# to the maximum number of cells found on any row

# Another comment and more rows following
|  # this is a row with two empty cells
```

AUTHOR
======

Tom Browder, `<tom.browder@gmail.com>` (`tbrowder` on IRC `#raku`)

COPYRIGHT & LICENSE
===================

Copyright (c) 2020 Tom Browder, all rights reserved.

This program is free software; you can redistribute it or modify
it under the same terms as Raku itself.

See that license [here](./LICENSE).
