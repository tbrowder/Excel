unit module Exec::Utility;

# A1 format utilities

sub split-ranges(%fmt, :$debug --> Hash) is export {
    my %new-fmt;

    # Cell "A1" keys like "B1-B6" and "B1-D1" are ranges and need to be split into
    # their own key/array
    KEY: for %fmt.keys.sort -> $k is copy {
        note "DEBUG-2: hjson key: '$k'" if $debug;
        my $v = %fmt{$k};
        my $vtyp = $v.^name;

        note "  its value type is '$vtyp'" if $debug;
        # if it's a range, split it, other just pass it on
        if $k ~~ /^ :i (<[A..Z]>+ <[1..9]> \d*) ['-' (<[A..Z]>+ <[1..9]> \d*) ]? $/ {
            say "DEBUG: key '$k' is an 'A1' key and should have an array or string as its value" if $debug;
            my $k1 = ~$0;
            my $k2;
            if defined $1 {
                $k2 = ~$1;
            }
            if !$k2.defined {
                %new-fmt{$k} = $v;
                next KEY;
            }

            # a cell range key: the values will be assigned to all the keys in the range
            my @ckeys = split-range $k;
            for @ckeys -> $ck {
                %new-fmt{$ck} = $v;
            }
            next KEY;
        }
        else {
            # not a cell key
            %new-fmt{$k} = $v;
        }
    }

    return %new-fmt;

} # from: sub split-ranges

sub split-range($range is copy, :$debug --> Array) is export {
    die "FATAL: Range '$range' has no hyphen" if !$range.contains: '-';
    my ($start, $end) = split '-', $range;

    {
        # this is mainly a debug check
        my ($alpha-start, $int-start, $alpha-end, $int-end);
        if $start ~~ /^ :i (<[A..Z]>+) (<[1..9]> \d*) $/ {
            $alpha-start = ~$0;
            $int-start   = +$1;
        }
        else {
            die "FATAL: Unexpected A1 format: '$start'";
        }

        if $end ~~ /^ :i (<[A..Z]>+) (<[1..9]> \d*) $/ {
            $alpha-end = ~$0;
            $int-end   = +$1;
        }
        else {
            die "FATAL: Unexpected A1 format: '$end'";
        }
    }

    my ($start-row, $start-col, $srow-abs, $scol-abs) = xl-cell-to-rowcol $start;
    my ($end-row  , $end-col  , $erow-abs, $ecol-abs) = xl-cell-to-rowcol $end;


    my @A1;
    if $start-row == $end-row {
        # column range, e.g., "A9:Z9"
        for $start-col .. $end-col -> $col {
            my $A1 = xl-rowcol-to-cell $start-row, $col;
            @A1.push: $A1;
        }
    }
    elsif $start-col == $end-col {
        # row range, e.g., "A1:A9"
        for $start-row .. $end-row -> $row {
            my $A1 = xl-rowcol-to-cell $row, $start-col;
            @A1.push: $A1;
        }
    }
    else {
        die "FATAL: unexpected non-linear cell range: '{$start}:{$end}'";
    }

    return @A1;

} # from: sub split-range


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

    # lowercase it


    return lc ($col-str ~ $row-abs ~ $row);

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
# The $row_absolute and $col_absolute parameters aren't documented
# because they mainly used internally and aren't very useful to the
# user.
#
sub xl-cell-to-rowcol($cell is copy, :$debug) is export {

    return (0, 0, 0, 0) unless $cell;
    note "DEBUG: input A1 \$cell = '$cell'" if 0 and $debug;

    # ensure we handle uppercase internally
    $cell .= uc;

    $cell ~~ /^ ('$'?) (<[A..Z]>**1..3) ('$'?) (\d+) /;

    my $col-abs = defined $0 ?? 1 !! 0;
    my $col     = ~$1;
    my $row-abs = defined $2 ?? 1 !! 0;
    my $row     = ~$3;

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
