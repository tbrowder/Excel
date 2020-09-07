[![Actions Status](https://github.com/tbrowder/Excel/workflows/test/badge.svg)](https://github.com/tbrowder/Excel/actions)

# Excel

**A Raku module to create or use Excel xlsx files**

# UPDATE 2020-04-28

**SEE WIP EXCEL TEMPLATE IN EXAMPLES DIRECTORY**

[Note this module replaces the short-lived module *Excel::Text::Template*.]

The module can:

* Read an existing Microsoft Excel xlsx file
* Create Excel xlsx files
* Use an Excel file as a template to create new Excel xlsx files
* Use an HJSON-format configuration file to input values defining a
  templated project
* Use an HJSON-format formats file to force  certain  Excel xlsx formatting
  that cannot currently be read from the Excell xlsx template file

Planned:

* Allow user-defined call-back functions to aid the templating
  process

## DESCRIPTION

This module provides the capability of using Excel templates
to generate Excel files. It is a WIP and has little working code
at the moment. If you are interested in the concept, please
star the project, follow it, and file a feature request issue.

Currently working code is in the "dev" directory.

The project uses several Perl modules which will have to
be installed for the distro to work (I use `cpanm` for that):

+ Perl modules required:

    + `Excel::Writer::XLSX`      # write files
    + `Spreadsheet::ParseXLSX`   # read files
    + `Spreadsheet::Read`        # read files
    + `Spreadsheet::Reader::ExcelXML`        # read files

## LIMITATIONS

Currently the reader is capable of extracting the following
from existing Excel xslx files:

+ Worksheet data:

  - name
  - cell type
  - cell formatted values
  - cell unformatted values
  - cell formulas

The reader is **not** capable of extracting format information
so the user must define any desired non-default output formatting
via an input HJSON file.

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

## PROVIDES

This version provides:

* Excel and CSV data readers

* Excel template reader

* Excel format reader

* Excel writer

    * a single workbook per data set

    * a single workbook with a single worksheet per data set

## Using a template

1. Manually create the Excel template to look as desired.
   Use dummy data in the desired cells and format.  Add real
   exlanatory text and format as desired.  Include working formulas
   using dummy data and format and locate result cells as desired.

2. Use special coded text inputs in the template cell to define mappings
   between input data and the final worksheet.

3. Use an HJSON file to define various output formats for a case if
   you want to transfer duplicate formatting in the template to the
   output worksheets.

CREDITS
=======

Many thanks to all the Perl and Raku authors whose modules I've used
over the last 25+ years, including all the well-known luminaries Larry
Wall and Damian Conway. But the workhorse modules I used most heavily
over 15 years in my civilian career were those by **John McNamara**,
the most recent incarnation of his great Excel modules being
`Excel::Writer::XLSX`.

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
