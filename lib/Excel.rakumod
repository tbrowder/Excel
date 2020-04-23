unit module Excel;

use Data::Dump;

=begin comment
# Perl modules
use File::LibMagic:from<Perl5>;
#use Excel::Writer::XLSX:from<Perl5>;
use Spreadsheet::XLSX:from<Perl5>;
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

sub parse-xlsx($fnam, :$wsnum = 0, :$wsnam, :$debug) is export {
    # Returns an array of rows which are arrays of columns of
    # cell values.
    my @rowcols = [];

    use Spreadsheet::ParseXLSX:from<Perl5>;
    my $parser = Spreadsheet::ParseXLSX.new;
    my $wb     = $parser.parse($fnam)
              || die "FATAL: File $fnam can't be parsed";
    note "DEBUG file: {$wb<File>}" if $debug;
    my $wsc = $wb.worksheet_count;
    note "DEBUG worksheet count: {$wsc}" if $debug;

    my $sn = 0;
    for 0..^$wsc -> $wsn {
        my $ws = $wb.worksheet($wsn); # can also use the name if need be
        if $sn && $debug {
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
        for $row-min ... $row-max -> $row {
            my @cols = [];
            for $col-min ... $col-max -> $col {
                my ($value, $unfmt, $equat) = '', '', '';
                my $cell  = $ws.get_cell($row, $col);
                unless $cell.defined {
                    $cell = '';
                    @cols.push: $value;
                    next;
                }
                $value = $cell.value;
                $unfmt = $cell.unformatted;
                #$equat = $cell.Formula;
                @cols.push: $value;
            }
            @rowcols.push: @cols;
        }
        ++$sn;
    }
    return @rowcols;
}

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

multi write-xlsx($fnam, $workbook) is export {
    # Writes an xlsx file as a copy of the input Excel workbook.
}
=end comment

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
            note "DEBUG: cell[$i][$j] = '$val'" if $debug;
            ++$j;
        }
        ++$i;
    }
    $wb.close;
}
