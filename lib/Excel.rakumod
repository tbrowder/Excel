unit module Excel;

use JSON::Hjson;
use Data::Dump::Tree;

use Excel::Format;
use Excel::Utility;

=begin pod

The following classes are used to capture XLSX workbook data to use
with the Perl XLSX writer C<Excel::Writer::XLSX>.

=end pod

class Cell is export {
    use Excel::Writer::XLSX:from<Perl5>;

    has $.row; # 0-indexed, defined in constructor
    has $.col; # 0-indexed, defined in constructor
    has $.A1;  # "A1" reference, defined in constructor

    has $.value       is rw = '';
    has $.unformatted is rw = '';
    has $.formula     is rw = '';
    has $.type        is rw = '';

    has $.coded-text  is rw = '';
    has $.debug       is rw = 0;

    #=begin comment
    has %.font        is rw = {}
    has $.format      is rw = ''; {}
    has $.encoding    is rw = '';
    has $.is_merged     is rw = '';
    has $.get_rich_text is rw = '';
    has $.get_hyperlink is rw = '';
    has %.properties    is rw = {}

    submethod TWEAK {
        # initialize known hash key/value pairs
        # font info
        self.font<Name>           = '';
        self.font<Bold>           = '';
        self.font<Italic>         = '';
        self.font<Height>         = '';
        self.font<Underline>      = '';
        self.font<UnderlineStyle> = '';
        self.font<Color>          = '';
        self.font<Strikeout>      = '';
        self.font<Super>          = '';

        # format info
        =begin comment
        self.format<Font>         = '';
        self.format<AlignH>       = '';
        self.format<AlignV>       = '';
        self.format<Indent>       = '';
        self.format<Wrap>         = '';
        self.format<Shrink>       = '';
        self.format<Rotate>       = '';
        self.format<JustLast>     = '';
        self.format<ReadDir>      = '';
        self.format<BdrStyle>     = '';
        self.format<BdrColor>     = '';
        self.format<BdrDiag>      = '';
        self.format<Fill>         = '';
        self.format<Lock>         = '';
        self.format<Hidden>       = '';
        self.format<Style>        = '';
        =end comment
    }

    method read-xlsx-cell($perl-ws, $perl-wc) {
        # Given Spreadsheet::ParseExcelXLSX worksheet and cell objects $ws, and $wc, read
        # their data into the calling Raku cell.

        # At the moment I see no reason not to transfer all
        # the known attritutes (properties).
        return if !$perl-wc;

        # fundamental
        self.value                = $perl-wc.value;

        if self.value ~~ /^';'/ {
            self.coded-text = self.value;
        }
        self.unformatted          = $perl-wc.unformatted;
        self.formula              = $perl-wc<Formula>;
    }

    #method write-xlsx-cell($perl-ws, $row, $col, :$format = '') {
    method write-xlsx-cell($perl-ws, $format = '') {
        # Given an Excel::Writer::XLSX worksheet object $ws, write
        # this Raku cell's attributes into the target cell at the same location.
        # We always use the A1 notation.

        # At the moment I see no reason not to transfer all
        # the known attritutes (properties).

        if self.debug {
            note "DEBUG: cell[{self.row}][{self.col}] is a Cell object";
            note "    A1          = '{self.A1}'";
            note "    value       = '{self.value}'";
            note "    unformatted = '{self.unformatted}'";
            note "    formula     = '{self.formula}'";
        }

        # now write to the real spreadsheet
        #my $i  = self.row;
        #my $j  = self.col;
        my $A1 = self.A1;
        my $written = 0;
        if self.formula {
            # we need A1 row/col ID
            #my $A1 = xl-rowcol-to-cell($i, $j);
            if $format {
                $perl-ws.write_formula: $A1, "{self.formula}", $format;
            }
            else {
                $perl-ws.write_formula: $A1, "{self.formula}";
            }
            ++$written;
        }
        if self.value {
            if $format {
                #$perl-ws.write_string: $i, $j, "{self.value}", $format;
                $perl-ws.write_string: $A1, "{self.value}", $format;
            }
            else {
                #$perl-ws.write_string: $i, $j, "{self.value}";
                $perl-ws.write_string: $A1, "{self.value}";
            }
            ++$written;
        }
        if self.unformatted {
            #$perl-ws.write: $i, $j, "{self.unformatted}";
            $perl-ws.write: $A1, "{self.unformatted}";
            ++$written;
        }

        unless $written {
            #$perl-ws.write_blank: $i, $j;
            $perl-ws.write_blank: $A1;
        }
    }

} # end of: class Cell


class Worksheet is export {
    has $.name;   # defined in constructor
    has $.number; # defined in constructor
    has @.rowcols is rw = [];
}

class Workbook is export {
    has $.filename; # defined in constructor
    has @.worksheets is rw = [];
}

sub parse-xlsx-workbook($filename, :$perl-wsnum = 0, :$perl-wsnam, :$debug) is export {
    # Returns a Raku copy of the ExcelXLSX workbook in the input file.

    use Spreadsheet::ParseXLSX:from<Perl5>;
    my $perl-parser = Spreadsheet::ParseXLSX.new;
    my $perl-wb     = $perl-parser.parse($filename)
              || die "FATAL: File $filename can't be parsed";
    note "DEBUG file: {$perl-wb<File>}" if $debug;
    my $perl-wsc = $perl-wb.worksheet_count;
    note "DEBUG worksheet count: {$perl-wsc}" if $debug;

    my $raku-wb = Workbook.new: :$filename;

    my $sn = 0;
    for 0..^$perl-wsc -> $perl-wsnum {
        my $perl-ws  = $perl-wb.worksheet($perl-wsnum); # can also use the name if need be
        my $perl-wsn = $perl-ws.get_name;

        # Raku
        my $raku-ws  = Worksheet.new: :number($perl-wsnum), :name($perl-wsn);
        $raku-wb.worksheets.push: $raku-ws;

        if 0 && $sn && $debug {
            note "DEBUG: exiting after first worksheet";
            exit;
        }
        if $debug {
            note "DEBUG: got Perl worksheet $sn...";
        }
        my ($row-min, $row-max) = $perl-ws.row_range;
        if $row-min > $row-max {
            die "FATAL: $row-min > $row-max";
        }
        if $debug {
            note "DEBUG: row min/max: {$row-min}/{$row-max}";
        }
        my ($col-min, $col-max) = $perl-ws.col_range;
        if $col-min > $col-max {
            die "FATAL: $col-min > $col-max";
        }

        ROW: for $row-min ... $row-max -> $row {
            my @cols = [];
            COL: for $col-min ... $col-max -> $col {
                # this is the Perl cell object from Spreadsheet::ParseXSLX:
                my $perl-wc  = $perl-ws.get_cell($row, $col);

                my $A1 = xl-rowcol-to-cell $row, $col;
                # capture it in a Raku object
                my $cell = Cell.new: :$row, :$col, :$A1;
                unless $perl-wc.defined {
                    $cell.value = '';
                    @cols.push: $cell;
                    next COL;
                }

                $cell.read-xlsx-cell: $perl-ws, $perl-wc;

                =begin comment
                $cell.value       = $wc.value       // '';
                $cell.unformatted = $wc.unformatted // '';
                $cell.formula     = $wc<Formula>    // '';
                =end comment

                # lots more data to collect. see Spreadsheet::ParseExcel


                # finished

                @cols.push: $cell;
            }
            $raku-ws.rowcols.push: @cols;
        }
        ++$sn;
    }

    return $raku-wb; # the Workbook

} # end of: sub parse-xlsx-workbook


sub write-xlsx-workbook($fnam,
                        Workbook $raku-wb,
                        :$hjfil,      #= optional Hjson file of formatting desires
                        :$debug,
                        :@ofils!,
                       ) is export {
    # Writes an xlsx file as a copy of the input Excel workbook.
    # lower case everything!!
    my %fmt = $hjfil ?? from-hjson(lc (slurp $hjfil)) !! {};

    use Excel::Writer::XLSX:from<Perl5>;

    # start an empty Excel file to be written to
    my $perl-wb  = Excel::Writer::XLSX.new: $fnam;
    # apply formatting as desired
    if %fmt.elems {
        # the formats are converted into format vars
        # inside the workbook for later use in cells
        build-xlsx-formats %fmt, $perl-wb, :$debug;
    }

    # iterate through the input workbook
    my @Wb-sheets = $raku-wb.worksheets;
    my $Wb-wsnums = $raku-wb.worksheets.elems;

    #ddt $Ws;

    my $k = -0;
    WORKSHEET: for @Wb-sheets -> $raku-ws {
        ++$k;

        my $nrows = $raku-ws.rowcols.elems;
        my $ncols = $raku-ws.rowcols[0].elems;
        note "DEBUG: writing $nrows rows and $ncols columns";

        my $perl-ws;
        if $raku-ws.name {
            $perl-ws  = $perl-wb.add_worksheet: "{$raku-ws.name}";
        }
        else {
            $perl-ws  = $perl-wb.add_worksheet;
        }

        my $perl-wsn = $perl-ws<Name> // '';
        note "Sheet name: $perl-wsn" if $debug;

        my $i = -1;
        ROW: for $raku-ws.rowcols -> $row {
            ++$i;

            my $j = -1;
            COL: for @($row) -> $cell {
                ++$j;

                # get the A1 name of the cell and its format, if any
                my $A1 = $cell.A1;
                my $format = %fmt{$A1}:exists && %fmt{$A1}.defined ?? %fmt{$A1} !! '';
                $cell.write-xlsx-cell: $perl-ws, $format;

                =begin comment
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
                =end comment

            } # end cell
        } # end row
    } # end worksheets

    $perl-wb.close;
    @ofils.push: $fnam;

} # end of: sub write-xlsx-workbook

sub build-xlsx-formats(%fmt,
                       $perl-wb,     #= a Perl Excel::Writer::XLSX workbook
                       :$debug) {
    # The formats are converted into format vars
    # inside the workbook for later use in cells
    # and returned as vars named for the cells.
    # For now, the formats are for all worksheets, but
    # could be named something like $B3-WS0 or $B3-WS'Foo for
    # specific sheets.

    use Excel::Writer::XLSX:from<Perl5>;

    # create eval strings for each format object to be defined
    # name them after cell, and possibly for specific worksheets
    my %e;
    # global attributes to be applied for all %e keys
    my %g;

    # format keys may be using ranges
    %fmt = split-ranges %fmt;

    if 0 && $debug {
        for %fmt.keys -> $k {
            ddt %fmt{$k};
        }
        die "DEBUG exit";
    }

    #my @keys = %fmt.keys.sort;
    KEY: for %fmt.keys.sort -> $k is copy {
        note "DEBUG-1: hjson key: '$k'" if $debug;
        my $v = %fmt{$k};
        my $vtyp = $v.^name;

        if $debug {
            note "  its value type is '$vtyp'";
            if $v ~~ Str {
                note "  its value is '$v'";
            }
        }

        # we need to handle a range of cells in the same row or column
        if $k ~~ /^ :i (<[A..Z]>+ <[1..9]> \d*) ['-' (<[A..Z]>+ <[1..9]> \d*) ]? $/ {
            say "DEBUG: key '$k' is an 'A1' key and should have an array or string as its value" if $debug;
            my $k1 = ~$0;
            my $k2;
            if defined $1 {
                $k2 = ~$1;
            }
            if $debug {
                if $k2 {
                    note "DEBUG: range key '$k', keys: '$k1' and '$k2'";
                }
                else {
                    note "DEBUG: non-range key '$k', key: '$k1'";
                }

                #note "  next key...";
                #next KEY;
            }

            # it must be an "A1" cell
            my $cell = $k;

            # Handle formatting: We could have a single value OR
            # an array of values.
            if $v ~~ Str {
                # could have a colon pair
                if $v ~~ /':'/ {
                    my ($a, $val) = split ':', $v;
                    die "FATAL: Format attr '$a' is not known." if not %formats{$a}:exists;
                    %e{$cell}{$a} = $val;
                }
                else {
                    %e{$cell}{$k} = $v;
                }

            }
            elsif $v ~~ Array {
                for $v -> $attr {
                    # could have a colon pair
                    if $attr ~~ /':'/ {
                        my ($a, $val) = split ':', $attr;
                        die "FATAL: Format attr '$a' is not known." if not %formats{$a}:exists;
                        %e{$cell}{$a} = $val;
                    }
                    else {
                        %e{$cell}{$attr} = Nil;
                    }
                }
            }
            else {
                die "FATAL: Unexpected value type '$vtyp' for key '$k'";
            }
        }
        else {
            # a global format attribute to be applied for the cell-named formats
            # is it a known format?
            die "FATAL: Format attr '$k' is not known." if not %formats{$k}:exists;
            %g{$k} = $v;
        }
    }

    if 0 and $debug {
        note "DEBUG early exit.";
        exit;
    }

    # apply all the global attrs to the cell attrs but only if they aren't already specified
    # in the border properties
    for %g.keys -> $gattr {
        my $gval = %g{$gattr};

        for %e.keys -> $cell {
            for %(%e{$cell}) -> $attr {
                # if it has a value we pass on
                next if %e{$cell}{$attr}.defined;

                # check: top, bottom, left, right for linewidth
                for "left", "right", "top", "bottom" -> $b {
                    if $b eq $attr {
                        %e{$cell}{$attr} = $gattr;
                    }
                }
            }
        }
    }
    # we now have all we need to define the formats
    for %e.keys -> $a1 {
        note "DEBUG: defining format for cell '$a1'" if $debug;
        my %prop = %(%e{$a1});
        for %prop.keys -> $p {
            my $v = %prop{$p}.defined ?? %prop{$p} !! '';
            given $p {
                when /border/ {
                }
                when /top/ {
                }
                when /bottom/ {
                }
                when /left/ {
                }
                when /right/ {
                }
                when /bold/ {
                }
                when /italic/ {
                }
                when /font/ {
                }
                when /color/ {
                }
                when /linewidth/ {
                }
                when /size/ {
                }
                when /^l$/ {
                }
                when /^c$/ {
                }
                when /^r$/ {
                }
                when /^fg$/ {
                }
                when /^bg$/ {
                }
            } # end of given block

        }
    }


} # end of: sub build-xlsx-formats
