[![Build Status](https://travis-ci.com/tbrowder/Excel-Raku.svg?branch=master)](https://travis-ci.com/tbrowder/Excel-Raku)

# Excel

**A Raku module to create or use Excel xlsx files**

# UPDATE 2020-04-18

**SEE WIP EXCEL TEMPLATE IN EXAMPLES DIRECTORY**

[Note this module replaces the short-lived module *Excel::Text::Template*.]

The module can:

* Read an existing Microsoft Excel xlsx file
* Create Excel xlsx files
* Use an Excel file as a template to create new Excel xlsx files

Planned:

* Use a text file as a template to create new Excel xlsx files
* Allow user-defined call-back functions to aid the templating
  process
* Use a TOML-format configuration file to input values defining a
  templated project

## DESCRIPTION

This module provides the capability of using Excel and text templates
to generate Excel files. It is a WIP and has little working code
at the moment. If you are interested in the concept, please
star the project, follow it, and file a feature request issue.

Currently working code is in the "dev" directory.

The project uses several Perl modules which will have to
be installed for the distro to work (I use `cpanm` for that):

+ Perl modules required:

    + `Excel::Writer::XLSX`
    + `Spreadsheet::XLSX`

## DATA FLOW

The major use case is designed to
take input consisting of one or more individual
data sets (such as the rows of a database,
CSV file, or spreadsheet) and convert
each row into an Excel workbook via
a template which describes the mapping
from input row columns to table cells
in the output Excel worksheet.

## USE CASES

Typical row-oriented data sets might be:

* a teacher's student list
* a manager's employee list 
* a research scientist's experimental results
* a financial analyst's list of security data

## INITIAL VERSION

The first working version will provide:

* Excel and data readers

* Excel and text template readers

* Excel writers 

    * a single workbook per data set

    * a single workbook with a single worksheet per data set

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
attribute pairs (using a syntax similar to `Raku`'s Pairs) following the cell content:

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

As I work on my real-world project (using an Excel template)
I realize the cell-mapping DSL needs to be a little more 
complex to be able to specify where the cell data are coming
from and where are they to be written. So far I am using
the first column of my template for "directives" that
are not written to the output file. Among the directives
I'm using are special row identifiers for grouping
data lines if needed as for my tax case where I
have multiple security buys for the same security over
the years, one line for each lot.

See the examples directory for my Excel template as it
progresse.

CREDITS
=======

Many thanks to all the Perl and Raku authors whose modules I've used
over the last 25+ years, including all the well-known luminaries
Larry Wall and Damian Conway. But the workhorse modules I used
most heavily over 15 years in my civilian career were those
by **John McNamara**, the most recent incarnation of his great
Excel modules being Excel::Writer::XLSX.

Of course I couldn't use John's work without the excellent
Raku module Inline::Perl5 whose original author is **Stefan Seifert**.

AUTHOR
======

Tom Browder, `<tom.browder@gmail.com>` (`tbrowder` on IRC `#raku`)

COPYRIGHT & LICENSE
===================

Copyright (c) 2020 Tom Browder, all rights reserved.

This program is free software; you can redistribute it or modify
it under the same terms as Raku itself.

See that license [here](./LICENSE).
