[![Build Status](https://travis-ci.com/tbrowder/Excel-Raku.svg?branch=master)](https://travis-ci.com/tbrowder/Excel-Raku)

# Excel

**A Raku module to create or use Excel xlsx files**

# UPDATE 2020-04-19

**SEE WIP EXCEL TEMPLATE IN EXAMPLES DIRECTORY**

[Note this module replaces the short-lived module *Excel::Text::Template*.]

The module can:

* Read an existing Microsoft Excel xlsx file
* Create Excel xlsx files
* Use an Excel file as a template to create new Excel xlsx files

Planned:

* Allow user-defined call-back functions to aid the templating
  process
* Use a TOML-format configuration file to input values defining a
  templated project

## DESCRIPTION

This module provides the capability of using Excel templates
to generate Excel files. It is a WIP and has little working code
at the moment. If you are interested in the concept, please
star the project, follow it, and file a feature request issue.

Currently working code is in the "dev" directory.

The project uses several Perl modules which will have to
be installed for the distro to work (I use `cpanm` for that):

+ Perl modules required:

    + `Excel::Writer::XLSX` # write files
    + `Spreadsheet::XLSX`   # read files

## DATA FLOW

The major use case is designed to take input consisting of one or more
individual data sets (such as the rows of a database, CSV file, or
spreadsheet) and convert each row into an Excel workbook via a
template which describes the mapping from input row columns to table
cells in the output Excel worksheet.

## USE CASES

Typical row-oriented data sets might be:

* a teacher's student list
* a manager's employee list
* a research scientist's experimental results
* a financial analyst's list of security data

## INITIAL VERSION

The first working version will provide:

* Excel and CSV data readers

* Excel template reader

* Excel writer

    * a single workbook per data set

    * a single workbook with a single worksheet per data set

## Using a template

After diving in to my project I've decided the original direction was
too complicated for my use. Now I'm heading toward a process like
this:

1. Design the Excel template to look just like I want it to look and
   use dummy data in the desired cells and format.  Add real
   exlanatory text and format as desired.  Include working formulas
   using dummy data and format and locate result cells as desired.

2. Use additional worksheets in the template to define mappings
   between the cells in the input, template, and output files.

## GENESIS

See the original README.md in the "old" subdir.

CREDITS
=======

Many thanks to all the Perl and Raku authors whose modules I've used
over the last 25+ years, including all the well-known luminaries Larry
Wall and Damian Conway. But the workhorse modules I used most heavily
over 15 years in my civilian career were those by **John McNamara**,
the most recent incarnation of his great Excel modules being
Excel::Writer::XLSX.

Of course I couldn't use John's work without the excellent Raku module
Inline::Perl5 whose original author is **Stefan Seifert**.

AUTHOR
======

Tom Browder, `<tom.browder@gmail.com>` (`tbrowder` on IRC `#raku`)

COPYRIGHT & LICENSE
===================

Copyright (c) 2020 Tom Browder, all rights reserved.

This program is free software; you can redistribute it or modify
it under the same terms as Raku itself.

See that license [here](./LICENSE).
