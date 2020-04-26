unit module Excel;

#use Data::Dump;
use Data::Dump::Tree;

=begin pod

The following classes are used to capture XLSX workbook data to use
with the Perl XLSX writer C<Excel::Writer::XLSX>.

=end pod


class Cell is export {
    has $.row; # defined in constructor
    has $.col; # defined in constructor
    has $.A1;  # defined in constructor
    has $.value       is rw = '';
    has $.unformatted is rw = '';
    has $.formula     is rw = '';
    has $.type        is rw = '';
    #has %.attrs       is rw = {};
}

class Worksheet is export {
    has $.name;   # defined in constructor
    has $.number; # defined in constructor
    has @.rowcols is rw = [];
    #has %.attrs   is rw = {};
}

class Workbook is export {
    has $.filename; # defined in constructor
    has @.worksheets is rw = [];
    #has %.attrs      is rw = {};
}

=begin comment
# Perl modules
use Excel::Writer::XLSX:from<Perl5>;
use Spreadsheet::ParseXLSX:from<Perl5>;
# older subs used with CVS::Parser
=end comment

=begin comment
sub copy-xlsx($fin, $fout, :$debug) is export {
    # This is mainly used to test the interoperability of
    # the reader and writer.
    my $wb-i = Spreadsheet::XLSX.new: $fin;
#    my $wb-o = Excel::Writer::XLSX.new: $fout;

    my @ws-i  = @($wb-i<Worksheet>);
    for @ws-i -> $ws {
        my $wsn = $ws<Name>;
        note "Sheet name: $wsn" if $debug;
    }
}
=end comment

sub parse-xlsx-workbook($filename, :$wsnum = 0, :$wsnam, :$debug) is export {
    # Returns a Raku copy of the ExcelXLSX workbook in the input file.
    #my @rowcols = [];

    use Spreadsheet::ParseXLSX:from<Perl5>;
    my $parser = Spreadsheet::ParseXLSX.new;
    my $wb     = $parser.parse($filename)
              || die "FATAL: File $filename can't be parsed";
    note "DEBUG file: {$wb<File>}" if $debug;
    my $wsc = $wb.worksheet_count;
    note "DEBUG worksheet count: {$wsc}" if $debug;

    my $Wb = Workbook.new: :$filename;

    my $sn = 0;
    for 0..^$wsc -> $wsnum {
        my $ws  = $wb.worksheet($wsnum); # can also use the name if need be
        my $wsn = $ws.get_name;
        my $Ws = Worksheet.new: :number($wsnum), :name($wsn);

        $Wb.worksheets.push: $Ws;

        if 0 && $sn && $debug {
            note "DEBUG: exiting after first worksheet";
            exit;
        }
        if $debug {
            note "DEBUG: got worksheet $sn...";
        }
        my ($row-min, $row-max) = $ws.row_range;
        if $debug {
            note "DEBUG: row min/max: {$row-min}/{$row-max}";
        }
        my ($col-min, $col-max) = $ws.col_range;

        ROW: for $row-min ... $row-max -> $row {
            my @cols = [];
            COL: for $col-min ... $col-max -> $col {
                # this is the Perl cell object from Spreadsheet::ParseXSLX:
                my $c  = $ws.get_cell($row, $col);
                my $A1 = xl-rowcol-to-cell $row, $col;

                # capture it in a Raku object
                my $cell = Cell.new: :$row, :$col, :$A1;
                unless $c.defined {
                    $cell.value = '';
                    @cols.push: $cell;
                    next COL;
                }
                $cell.value       = $c.value       // '';
                $cell.unformatted = $c.unformatted // '';
                $cell.formula     = $c<Formula>    // '';
       
                # lots more data to collect. see Spreedsheet::ParseExcel


                # finished

                @cols.push: $cell;
            }
            $Ws.rowcols.push: @cols;
        }
        ++$sn;
    }

    return $Wb; # the Workbook

    #return @rowcols;
}

=begin comment
sub read-xlsx($fnam, :$wsnum = 0, :$wsnam, :$debug) is export {
    # Returns an array of rows which are arrays of columns of
    # cell values.
    my @rowcols;

    use Spreadsheet::XLSX:from<Perl5>;
    my $wb = Spreadsheet::XLSX.new: $fnam;

    for $wb.keys.sort -> $k {
        note "DEBUG: found \$wb key: $k";
    }

    my $ws;
    if $wsnam {
        # a hack
        $ws = get-worksheet-by-name $wb, $wsnam;
    }
    else {
        $ws = $wb<Worksheet>[$wsnum];
    }
    for $ws.keys.sort -> $k {
        note $k.^name;
        note "DEBUG: found \$ws key: $k";
    }

    my $wsn = $ws<Name>;
    # name could be a Blob
    if $wsn ~~ Blob {
        $wsn .= decode;
    }

    note "Sheet name: $wsn" if $debug;
    for 0 .. $ws<MaxRow> -> $row {
        my @cols = []; # makes it an array, a single object
        for 0 .. $ws<MaxCol> -> $col {
            my $cell = $ws<Cells>[$row][$col];
            #Dump({$cell}) if $debug;
            #dd $cell if $debug;
            my $val  = $cell<Val> // ''; # Nil?
            # could be a Blob
            if $val ~~ Blob {
                $val .= decode;
            }
            my $typ  = $cell<Type> // ''; # Nil?
            # could be a Blob
            if $typ ~~ Blob {
                $typ .= decode;
            }
            my ($value, $unfmt, $type) = '', '', '';
            if $val {
                $value = $cell.value;
                # could be a Blob
                if $value ~~ Blob {
                    $value .= decode;
                }
                $unfmt = $cell.unformatted;
                # could be a Blob
                if $unfmt ~~ Blob {
                    $unfmt .= decode;
                }
                $type = $cell.unformatted;
                # could be a Blob
                if $type ~~ Blob {
                    $type .= decode;
                }
            }
            if $debug {
                note "row $row; col $col:";
                note "    Val:         |$val|";
                note "    Type:        |$typ|";
                note "    value:       |$value|";
                note "    unformatted: |$unfmt|";
            }
            @cols.push: $val;
        }
        @rowcols.push: @cols;
    }
    # NOTE: DO NOT TRY TO CLOSE THE WORKBOOK
    #       OR IT WILL CAUSE AN EXCEPTION
    #$wb.close;

    return @rowcols;

    sub get-worksheet-by-name($workbook, $wsnam) {
        my $n = 0;
        my @ws;
        for $workbook<Worksheet> -> $ws {
            my $wsn = $ws<Name>;
            return $ws if $wsn eq $wsnam;
            # record the info for possible exit
            my $s = sprintf "Number: %3d ; name: $wsn", $n++;
            @ws.push: $s;
        }
        # We gracefully exit here but show the worksheet names
        # and numbers that do exist:
        note "FATAL: No worksheet name '$wsnam' found:";
        .note for @ws;
        exit;
    }
}
=end comment

=begin comment
multi write-xlsx($fnam, @rows, :$debug) is export {
    # Writes an xlsx file from an input 2x2 array of xlsx cell
    # objects.

    # start an empty Excel file to be written to
    my $wb  = Excel::Writer::XLSX.new: $fnam;
    my $ws  = $wb<Worksheet>[0];
    my $wsn = $ws<Name>;
    note "Sheet name: $wsn" if $debug;
    my $nrows = @rows.elems;
    my $ncols = @rows[0].elems;
    for 0 .. ^$nrows -> $row {
        for 0 .. ^$ncols -> $col {
            my $in-cell = @rows[$row][$col];
            # test the object to ensure it's a Cell object

            =begin comment
            my $cell = $ws<Cells>[$row][$col];
            my $val  = $cell<Val>; # // ''; # Nil?
            note "row $row; col $col: |$val|" if $debug;
            if $values {
                @cols.push: $val;
            }
            else {
                @cols.push: $cell;
            }
            =end comment
        }
        #@rows.push: @cols;
    }
    #$wb.close;

    #return @rows;
}

=end comment

sub write-xlsx-workbook($fnam, Workbook $Wb, :$debug) is export {
    # Writes an xlsx file as a copy of the input Excel workbook.

    use Excel::Writer::XLSX:from<Perl5>;
    #use Excel::Writer::XLSX::Utility:from<Perl5>;

    # start an empty Excel file to be written to
    my $wb  = Excel::Writer::XLSX.new: $fnam;

    # iterate through the input workbook
    my @Wb-sheets = $Wb.worksheets;
    my $Wb-wsnums = $Wb.worksheets.elems;

    #ddt $Ws;

    my $k = 0;
    WORKSHEET: for @Wb-sheets -> $Ws {
        my $nrows = $Ws.rowcols.elems;
        my $ncols = $Ws.rowcols[0].elems;
        note "DEBUG: writing $nrows rows and $ncols columns";

        my $ws  = $wb.add_worksheet: {$Ws.name}; 

        my $wsn = $ws<Name> // '';
        note "Sheet name: $wsn" if $debug;

        my $i = 0;
        ROW: for $Ws.rowcols -> $row {
            my $j = 0;
            COL: for $row -> $cell {
                #my $cell  = $Ws.rowcols[$i][$j] // '';

                if !$cell {
                    $ws.write_blank: $i, $j;
                    next COL;
                }
                if !$ws {
                    die "Unexpected null Worksheek";
                }

                my $equat = $cell.formula     // '';
                my $value = $cell.value       // '';
                my $unfmt = $cell.unformatted // '';

                if $debug {
                    note "DEBUG: cell[$i][$j] is a Cell object";
                    note "    value       = '$value'";
                    note "    unformatted = '$unfmt'";
                    note "    formula     = '$equat'";
                }

                # now write to the real spreadsheet
                if $equat {
                    # we need A1 row/col ID
                    my $A1 = xl-rowcol-to-cell($i, $j);    # C2                $ws.write: $i, $j, $equat;
                    $ws.write_formula: $A1, $equat;
                }
                elsif $value {
                    $ws.write_string: $i, $j, $value;
                }
                elsif $unfmt {
                    $ws.write: $i, $j, $unfmt;
                }
                else {
                    $ws.write_blank: $i, $j;
                }
                ++$j;
            }
            ++$i;
        } # end row
        ++$k;
    } # end worksheets

    $wb.close;
}

sub write-xlsx-cells($fnam, @rowcols, :$debug) is export {
    # Writes an xlsx file from an input 2x2 array of Cell objects.

    use Excel::Writer::XLSX:from<Perl5>;
    use Excel::Writer::XLSX::Utility:from<Perl5>;

    # start an empty Excel file to be written to
    my $wb  = Excel::Writer::XLSX.new: $fnam;
    my $ws  = $wb.add_worksheet; # $wb<Worksheet>[0];
    my $wsn = $ws<Name> // '';
    note "Sheet name: $wsn" if $debug;

    my $nrows = @rowcols.elems;
    my $ncols = @rowcols[0].elems;
    my $i = 0;
    ROW: for @rowcols -> $row {
        my $j = 0;
        COL: for $row -> $col {
            my $cell  = @rowcols[$i][$j] // '';
            if !$cell {
                $ws.write_blank: $i, $j;
                next COL;
            }
            if !$ws {
                die "Unexpected null Worksheek";
            }

            my $equat = $cell.formula     // '';
            my $value = $cell.value       // '';
            my $unfmt = $cell.unformatted // '';

            if $debug {
                note "DEBUG: cell[$i][$j] is a Cell object";
                note "    value       = '$value'";
                note "    unformatted = '$unfmt'";
                note "    formula     = '$equat'";
            }

            # now write to the real spreadsheet
            if $equat {
                # we need A1 row/col ID
                my $A1 = xl-rowcol-to-cell($i, $j);    # C2                $ws.write: $i, $j, $equat;
                $ws.write_formula: $A1, $equat;
            }
            elsif $value {
                $ws.write_string: $i, $j, $value;
            }
            elsif $unfmt {
                $ws.write: $i, $j, $unfmt;
            }
            else {
                $ws.write_blank: $i, $j;
            }
            ++$j;
        }
        ++$i;
    }
    $wb.close;
}

sub write-xlsx-values($fnam, @rowcols, :$debug) is export {
    # Writes an xlsx file from an input 2x2 array of cells of text or
    # numerical data.

    use Excel::Writer::XLSX:from<Perl5>;
    # start an empty Excel file to be written to
    my $wb  = Excel::Writer::XLSX.new: $fnam;
    my $ws  = $wb<Worksheet>[0];
    my $wsn = $ws<Name> // '';
    note "Sheet name: $wsn" if $debug;

    my $nrows = @rowcols.elems;
    my $ncols = @rowcols[0].elems;
    my $i = 0;
    for @rowcols -> $row {
        my $j = 0;
        for $row -> $col {
            my $val = @rowcols[$i][$j];
            note "DEBUG: cell[$i][$j] contents = '$val'" if $debug;
            ++$j;
        }
        ++$i;
    }
    $wb.close;
}

##### Functions ported from Excel::Writer::XLSX
###############################################################################
#
# xl_rowcol_to_cell($row, $col, $row_absolute, $col_absolute)
#
sub xl-rowcol-to-cell($row is copy,
                      $col,
                      $row-abs is copy = 0,
                      $col-abs is copy = 0;
                     ) is export {

    ++$row;  # Change from 0-indexed to 1 indexed.
    $row-abs = $row-abs ?? '$' !! '';
    $col-abs = $col-abs ?? '$' !! '';

    my $col-str = xl-col-to-name($col, $col-abs);

    return $col-str ~ $row-abs ~ $row;

    =begin comment
    # original
    my $row     = $_[0] + 1;          # Change from 0-indexed to 1 indexed.
    my $col     = $_[1];
    my $row_abs = $_[2] ? '$' : '';
    my $col_abs = $_[3] ? '$' : '';


    my $col_str = xl_col_to_name( $col, $col_abs );

    return $col_str . $row_abs . $row;
    =end comment
}

###############################################################################
#
# xl_cell_to_rowcol($string)
#
# Returns: ($row, $col, $row_absolute, $col_absolute)
#
# The $row_absolute and $col_absolute parameters aren't documented because they
# mainly used internally and aren't very useful to the user.
#
sub xl-cell-to-rowcol($cell is copy) is export {

    return (0, 0, 0, 0) unless $cell;

    $cell ~~ / ('$'?) ([A..Z]**1..3) ('$'?)(\d+) /;

    my $col-abs = $0 eq "" ?? 0 !! 1;
    my $col     = $1;
    my $row-abs = $2 eq "" ?? 0 !! 1;
    my $row     = $3;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars = $col.comb;
    my $expn = 0;
    $col = 0;

    while @chars {
        my $char = @chars.pop;    # LS char first
        $col += ( ord( $char ) - ord( 'A' ) + 1 ) * ( 26**$expn );
        ++$expn;
    }

    # Convert 1-index to zero-index
    --$row;
    --$col;

    return $row, $col, $row-abs, $col-abs;

    =begin comment
    # original
    my $cell = shift;

    return ( 0, 0, 0, 0 ) unless $cell;

    $cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/;

    my $col_abs = $1 eq "" ? 0 : 1;
    my $col     = $2;
    my $row_abs = $3 eq "" ? 0 : 1;
    my $row     = $4;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars = split //, $col;
    my $expn = 0;
    $col = 0;

    while ( @chars ) {
        my $char = pop( @chars );    # LS char first
        $col += ( ord( $char ) - ord( 'A' ) + 1 ) * ( 26**$expn );
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col, $row_abs, $col_abs;
    =end comment
}

###############################################################################
#
# xl_col_to_name($col, $col_absolute)
#
sub xl-col-to-name($col is copy, $col-abs is copy) {

    $col-abs    = $col-abs ?? '$' !! '';
    my $col-str = '';

    # Change from 0-indexed to 1 indexed.
    ++$col;

    while $col {

        # Set remainder from 1 .. 26
        my $remainder = $col % 26 || 26;

        # Convert the $remainder to a character. C-ishly.
        my $col-letter = chr( ord( 'A' ) + $remainder - 1 );

        # Accumulate the column letters, right to left.
        $col-str = $col-letter ~ $col-str;

        # Get the next order of magnitude.
        $col = Int( ( $col - 1 ) div 26 );
    }

    return $col-abs ~ $col-str;

    =begin comment
    # original
    my $col     = $_[0];
    my $col_abs = $_[1] ? '$' : '';
    my $col_str = '';

    # Change from 0-indexed to 1 indexed.
    $col++;

    while ( $col ) {

        # Set remainder from 1 .. 26
        my $remainder = $col % 26 || 26;

        # Convert the $remainder to a character. C-ishly.
        my $col_letter = chr( ord( 'A' ) + $remainder - 1 );

        # Accumulate the column letters, right to left.
        $col_str = $col_letter . $col_str;

        # Get the next order of magnitude.
        $col = int( ( $col - 1 ) / 26 );
    }

    return $col_abs . $col_str;
    =end comment
}
