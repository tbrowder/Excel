unit module Excel;

use JSON::Hjson;
#use Data::Dump::Tree;

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
        =begin comment
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
    # another hash for format objects
    my %perl-fmt;

    # handle the input format data

    if %fmt.elems {
        # the formats are converted into format vars
        # inside the workbook for later use in cells
        build-xlsx-formats %fmt, $perl-wb, :%perl-fmt, :$debug;
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
        note "DEBUG: writing $nrows rows and $ncols columns" if $debug;

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
                my $format = %perl-fmt{$A1}:exists && %perl-fmt{$A1} ?? %perl-fmt{$A1} !! '';
                $cell.write-xlsx-cell: $perl-ws, $format;

            } # end cell
        } # end row
    } # end worksheets

    $perl-wb.close;
    @ofils.push: $fnam;

} # end of: sub write-xlsx-workbook

sub build-xlsx-formats(%fmt,
                       $perl-wb,     #= a Perl Excel::Writer::XLSX workbook
                       :%perl-fmt!,  #= stash format objects here keyed by cell A1 reference
                       :$debug) {
    # The formats are converted into format vars
    # inside the workbook for later use in cells
    # and returned as vars named for the cells (e.g., $D4).
    # For now, the formats are for all worksheets, but
    # could be named something like $B3-WS0 or $B3-WS'Foo for
    # specific sheets.

    use Excel::Writer::XLSX:from<Perl5>;

    # cell attributes
    my %e;
    # global attributes to be applied for all %e keys
    my %g;

    # format keys may be using ranges or groups
    %fmt = split-ranges %fmt;

    if 0 && $debug {
        for %fmt.keys -> $k {
            #ddt %fmt{$k};
        }
        die "DEBUG exit";
    }

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

        if $k ~~ /^ :i <cell> $/ {
            say "DEBUG: key '$k' is an 'A1' key and should have an array or string as its value" if $debug;
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
                for @($v) -> $attr {
                    # could have a colon pair
                    if $attr ~~ /':'/ {
                        my ($a, $val) = split ':', $attr;
                        die "FATAL: Format attr '$a' is not known." if not %formats{$a}:exists;
                        %e{$cell}{$a} = $val;
                    }
                    else {
                        %e{$cell}{$attr} = '';
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


        # the $cell value is the "A1"
        for %e.keys -> $cell {
            for %(%e{$cell}) -> $attr {
                note "DEBUG: cell attr type {$attr.^name}" if 0;
                # if it has a value we pass on to the next attr
                note "DEBUG: cell $cell has attr $attr" if 0;
                next if %e{$cell}{$attr}; #.defined;

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
        my $format = $perl-wb.add_format;

        # save that object in our %perl-fmt hash
        %perl-fmt{$a1}<perl-wb-fmt> = $format;

        # add to the format the properties we discovered
        my %prop = %(%e{$a1});
        for %prop.keys -> $p {
            my $v = %prop{$p}.defined ?? %prop{$p} !! '';
            given $p {
                note "DEBUG: handling property: $p";
                # border properties: 0-13
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
                when /border_color/ {
                }
                when /bottom_color/ {
                }
                when /top_color/ {
                }
                when /left_color/ {
                }
                when /right_color/ {
                }

                # font properties
                =begin comment
                when /font/ {
                    # TODO this causes coredump:
                    my $f = 'Courier';
                    #note "DEBUG: 'set_font' value: '$f'";
                    #note "DEBUG: 'set_font' value: '$v'";
                    #$format.set_font: $v;
                    #$format.set_font: $f;
                }
                when /size/ {
                    # TODO this causes coredump:
                    #note "DEBUG: 'set_size' value: $v";
                    #$format.set_size: $v;
                }
                when /bold/ {
                    $format.set_bold;
                }
                when /italic/ {
                    $format.set_italic;
                }
                when /underline/ {
                    $format.set_underline;
                }
                when /color/ {
                    $format.set_color: $v;
                }
                when /linewidth/ {
                    $format.set_linewidth: $v;
                }
                when /^l$/ {
                    $format.set_align: "left";
                }
                when /^c$/ {
                    $format.set_align: "center";
                }
                when /^r$/ {
                    $format.set_align: "right";
                }
                when /^fg$/ {
                    $format.set_fg_color: $v;
                }
                when /^bg$/ {
                    $format.set_bg_color: $v;
                }
                =end comment
            } # end of given block

        }
    }


} # end of: sub build-xlsx-formats
