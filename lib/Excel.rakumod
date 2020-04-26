unit module Excel;

use Data::Dump::Tree;

=begin pod

The following classes are used to capture XLSX workbook data to use
with the Perl XLSX writer C<Excel::Writer::XLSX>.

=end pod


class Cell is export {
    use Excel::Writer::XLSX:from<Perl5>;

    has $.row; # defined in constructor
    has $.col; # defined in constructor
    has $.A1;  # defined in constructor

    has $.value       is rw = '';
    has $.unformatted is rw = '';
    has $.formula     is rw = '';
    has $.type        is rw = '';

    has $.debug       is rw = 0;

    has %.font        is rw = {}
    has %.format      is rw = {}

    method TWEAK {
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
    }

    method read-xlsx-cell($ws, $wc) {
        # Given Excel::Writer::XLSX worksheet and cell objects $ws, and $wc, read
        # their data into the calling Raku cell.

        # At the moment I see no reason not to transfer all
        # the known attritutes (properties).
        return if !$wc;

        # fundamental
        self.value                = $wc.value;
        self.unformatted          = $wc.unformatted;
        self.formula              = $wc<Formula>;

        # font info
        self.font<Name>           = $wc<Font><Name>;
        self.font<Bold>           = $wc<Font><Bold>;
        self.font<Italic>         = $wc<Font><Italic>;
        self.font<Height>         = $wc<Font><Height>;
        self.font<Underline>      = $wc<Font><Underline>;
        self.font<UnderlineStyle> = $wc<Font><UnderlineStyle>;
        self.font<Color>          = $wc<Font><Color>;
        self.font<Strikeout>      = $wc<Font><Strikeout>;
        self.font<Super>          = $wc<Font><Super>;

        # format info
        self.format<Font>         = $wc<Format><Font>;
        self.format<AlignH>       = $wc<Format><AlignH>;
        self.format<AlignV>       = $wc<Format><AlignV>;
        self.format<Indent>       = $wc<Format><Indent>;
        self.format<Wrap>         = $wc<Format><Wrap>;
        self.format<Shrink>       = $wc<Format><Shrink>;
        self.format<Rotate>       = $wc<Format><Rotate>;
        self.format<JustLast>     = $wc<Format><JustLast>;
        self.format<ReadDir>      = $wc<Format><ReadDir>;
        self.format<BdrStyle>     = $wc<Format><BdrStyle>;
        self.format<BdrColor>     = $wc<Format><BdrColor>;
        self.format<BdrDiag>      = $wc<Format><BdrDiag>;
        self.format<Fill>         = $wc<Format><Fill>;
        self.format<Lock>         = $wc<Format><Lock>;
        self.format<Hidden>       = $wc<Format><Hidden>;
        self.format<Style>        = $wc<Format><Style>;
    }

    method write-xlsx-cell($ws, $row, $col) {
        # Given an Excel::Writer::XLSX worksheet object $ws, write
        # this Raku cell's attributes into the target cell.

        # At the moment I see no reason not to transfer all
        # the known attritutes (properties).

        if self.debug {
            note "DEBUG: cell[{self.row}][{self.col}] is a Cell object";
            note "    value       = '{self.value}'";
            note "    unformatted = '{self.unformatted}'";
            note "    formula     = '{self.formula}'";
        }

        # now write to the real spreadsheet
        my $i = self.row;
        my $j = self.col;
        my $written = 0;
        if self.formula {
            # we need A1 row/col ID
            my $A1 = xl-rowcol-to-cell($i, $j);
            $ws.write_formula: $A1, "{self.formula}";
            ++$written;
        }
        if self.value {
            $ws.write_string: $i, $j, "{self.value}";
            ++$written;
        }
        if self.unformatted {
            $ws.write: $i, $j, "{self.unformatted}";
            ++$written;
        }

        # formatting

        unless $written {
            $ws.write_blank: $i, $j;
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
        my $Ws  = Worksheet.new: :number($wsnum), :name($wsn);

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
                my $wc  = $ws.get_cell($row, $col);
                my $A1 = xl-rowcol-to-cell $row, $col;

                # capture it in a Raku object
                my $cell = Cell.new: :$row, :$col, :$A1;

                $cell.read-xlsx-cell: $ws, $wc;

                =begin comment
                unless $c.defined {
                    $cell.value = '';
                    @cols.push: $cell;
                    next COL;
                }
                $cell.value       = $wc.value       // '';
                $cell.unformatted = $wc.unformatted // '';
                $cell.formula     = $wc<Formula>    // '';
                =end comment

                # lots more data to collect. see Spreadsheet::ParseExcel


                # finished

                @cols.push: $cell;
            }
            $Ws.rowcols.push: @cols;
        }
        ++$sn;
    }

    return $Wb; # the Workbook

} # end of: sub parse-xlsx-workbook


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

    my $k = -0;
    WORKSHEET: for @Wb-sheets -> $Ws {
        ++$k;

        my $nrows = $Ws.rowcols.elems;
        my $ncols = $Ws.rowcols[0].elems;
        note "DEBUG: writing $nrows rows and $ncols columns";

        my $ws;
        if $Ws.name {
            $ws  = $wb.add_worksheet: "{$Ws.name}";
        }
        else {
            $ws  = $wb.add_worksheet;
        }

        my $wsn = $ws<Name> // '';
        note "Sheet name: $wsn" if $debug;

        my $i = -1;
        ROW: for $Ws.rowcols -> $row {
            ++$i;

            my $j = -1;
            COL: for @($row) -> $cell {
                ++$j;

                $cell.write-xlsx-cell: $ws, $i, $j;

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

    $wb.close;

} # end of: sub write-xlsx-workbook

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

} # end of: sub xl-rowcol-to-cell

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

} # end of sub: sub xl-cell-to-rowcol

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

} # end of: sub xl-col-to-name
