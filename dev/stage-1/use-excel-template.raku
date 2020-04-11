#!/usr/bin/env raku

use Text::Utils :normalize-string :strip-comment;

use Excel::Writer::XLSX:from<Perl5>;
use Spreadsheet::XLSX:from<Perl5>;

#use lib <./lib>;
#use Grammar;

#class Grammar {...}

my $ifil = 'hand-generated-template.xlsx';
my $ofil = '';
my $prog = $*PROGRAM.basename;
if !@*ARGS.elems {
    say qq:to/HERE/;
    Usage: $prog go | -i=<template file name> [-o=<output-file.xslx>]

    Creates an Excel xlsx file from the input template file.

    The default input file name is: '$ifil' and the default
    output file name is the input file name with any existing
    suffix starting with a '.' being replaced by '.xlsx'.

    For example, "template.txt" becomes "template.xlsx" upon success.

    If the input file is an xlsx file and no output is specified,
    then the file name gets a gets a slightly more complex change, e.g.,
    input file 'input.xlsx' becomes output file 'input-xlxs.xlsx'.
    HERE
    exit;
}

my $debug = 0;
for @*ARGS {
    when /^ '-i=' (\S*) $/ {
        $ifil = ~$0;
    }
    when /^ '-o=' (\S*) $/ {
        $ofil = ~$0;
    }
    when /^ d $/ {
        $debug = 1;
    }
}

# the input name must be *.xlsx for this test
if $ifil !~~ /'.xlsx' $/ {
    note "FATAL: The input file MUST have a '.xlsx' extension and be a valid Excel file.";
    exit;
}

# output file naming problems
if !$ofil {
    my $base = $ifil;
    $base ~~ s/'.' \N* $//;
    $base ~= '-xlsx.xlsx';
    $ofil = $base;
}
elsif $ofil eq $ifil {
    note "FATAL: The input and outout filenames are the same.";
    exit;

}
elsif $ofil !~~ /'.xlsx' $/ {
    note "FATAL: The output file MUST have a '.xlsx' extension and be a valid Excel file.";
    exit;
}

if $debug {
    note "DEBUG: in '$ifil'; out '$ofil'...exiting";
    exit;
}

# initial flow test
# read input
my $wb = Spreadsheet::XLSX.new: $ifil;
my $ws = @($wb<Worksheet>)[0];
my $wsn = $ws<Name>;
say "Sheet name: $wsn";
$ws<MaxRow> ||= $ws<MinRow>;
for $ws<MinRow> .. $ws<MaxRow> -> $row {
    $ws<MaxCol> ||= $ws<MinCol>;
    for $ws<MinCol> .. $ws<MaxCol> -> $col {
        my $cell = $ws<Cells>[$row][$col];
        if $cell {
            say "($row, $col) => {$cell<Val>}";
        }
    }
}

