###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use Excel;

use Test;

plan 4;

###############################################################################
#
# Tests setup.
#
my $got;
my @got;
my $expected;
my @expected;
my $caption;
my $cell;

# Create a test case for a range of the Excel 2007 columns.
$cell = 'A';
for 0 .. 300 -> $i {
    push @expected, [ $i, $i, $cell ~ ( $i + 1 ) ];
    ++$cell;
}

$cell = 'WQK';
for 16_000 .. 16_384 -> $i {
    push @expected, [ $i, $i, $cell ~ ( $i + 1 ) ];
    ++$cell;
}

###############################################################################
#
# Test the xl_rowcol_to_cell method.
#
$caption = " \tUtility: xl-rowcol-to-cell()";

for @expected -> $exp {
    push @got,
      [ $exp[0], $exp[1], xl-rowcol-to-cell($exp[0], $exp[1]) ];
}

is-deeply @got, @expected, $caption;

=begin comment
# original
for my $aref ( @$expected ) {
    push @$got,
      [ $aref->[0], $aref->[1], xl_rowcol_to_cell( $aref->[0], $aref->[1] ) ];
}

is_deeply( $got, $expected, $caption );
=end comment

###############################################################################
#
# Test the xl_rowcol_to_cell method with absolute references.
#
$expected = 'A$1';
$got = xl-rowcol-to-cell(0, 0, 1);
is $got, $expected, $caption;

$expected = '$A1';
$got = xl-rowcol-to-cell(0, 0, 0, 1);
is $got, $expected, $caption;


$expected = '$A$1';
$got = xl-rowcol-to-cell(0, 0, 1, 1);
is $got, $expected, $caption;
