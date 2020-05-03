unit module Exec::Utility;

# A1 format utilities

# some useful regexes
# use them as named captures like this:
#   if $str ~~ /^ <cell> $/ {
#       my $key = ~$/;
#   }

our token cell       is export { :i <[A..Z]>+ <[1..9]> \d* }               # no hyphens
our token line-range is export { <cell> '-' <cell> }                       # one hyphen
our token rect-range is export { <cell> '-' <cell> '-' <cell> '-' <cell> } # three hyphens
our token cell-group is export { <cell> [ <[,\s]>+ <cell> ]+ }             # a group of two or more cells

# a utility class for local use
class C {
    has Str $.A1; # is rw; # A1..ZZZ999 # MaxColsMaxRows
    # use these for sorting:
    has Int $.r ; # is rw; # 1..MaxRows
    has Str $.c ; # is rw; # A..MaxCols

    # 0-indexed rowcol
    has $.row;
    has $.col;

    submethod TWEAK {
        if $!A1 ~~ /^ (<[a..z]>+) (<[1..9]> \d*) $/ {
            $!c = ~$0;
            $!r = +$1;
            ($!row, $!col) = xl-cell-to-rowcol $!A1;
        }
        else {
            die "FATAL: Unexpected format (must use lower-case) in cell A1 label '$!A1'";
        }
     }
     
     sub sort-cols(C @c) {
     }

     sub sort-rows(C @c) {
     }
}

sub split-ranges(%fmt, :$debug --> Hash) is export {
    my %new-fmt;

    # Cell "A1" keys like "B1-B6" and "B1-B4-D1-D4" are ranges and
    # need to be split into their own key/array.
    # We also handle groups.
    KEY: for %fmt.keys.sort -> $k is copy {
        my $is-cell-key = 0;
        my $is-cell-grp = 0;

        note "DEBUG-2: hjson key: '$k'" if $debug;
        my $v = %fmt{$k};
        my $vtyp = $v.^name;

        my @k;
        if $k ~~ /^ <rect-range> $/ {
            my $c = ~$/;
            @k = split '-', $c;
            ++$is-cell-key;
        }
        elsif $k ~~ /^ <line-range> $/ {
            my $c = ~$/;
            @k = split '-', $c;
            ++$is-cell-key;
        }
        elsif $k ~~ /^ <cell-group> $/ {
            my $c = ~$/;
            # convert commas to spaces
            $c ~~ s:g/','/ /;
            @k = $c.words;
            ++$is-cell-key;
            ++$is-cell-grp;
        }
        elsif $k ~~ /^ <cell> $/ {
            my $c = ~$/;
            @k.push: $c;
            ++$is-cell-key;
        }

        note "  its value type is '$vtyp'" if $debug;
        # if it's a cell key or range, split it, otherwise just pass it on
        if $is-cell-key  {
            my $ncells = @k.elems;

            say "DEBUG: we have one or more 'A1' cell keys and should have an array or string as its shared value" if $debug;

            if $ncells == 1 {
                %new-fmt{$k} = $v;
                next KEY;
            }

            # a cell group key: the values will be assigned to all the
            # keys in the group
            if $is-cell-grp {
                # the cells have already been split out above
                for @k -> $k {
                    %new-fmt{$k} = $v;
                }
                next KEY;
            }

            # a cell range key: the values will be assigned to all the
            # keys in the range
            my @ckeys = split-range @k, :$debug;
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

} # end of: sub split-ranges

sub split-linear(@bcells, :$debug --> Array) is export {
    # the two @bcells are the bounding cells of a linear range
    die "FATAL: there should be two cells but there are {@bcells.elems}" if @bcells.elems != 2;

    # if there are two cells (one hyphen, linear range) it's easy:
    my @c;
    for @bcells -> $A1 {
        # if we have been rigorous in our plan the alpha chars
        # should be lower-case
        my $c = C.new: :$A1;
        if $debug {
            note "DEBUG: linear range, cell.A1: {$c.A1}";
        }
        @c.push: $c;
    }

    my ($L, $R, $T, $B); # range end cells: left/right (row range), top/bottom (column range)

    enum RangeStat <IsRow IsCol>;
    my $range-type;
    if @c.head.r == @c.tail.r {
        $range-type = IsRow;
        $L = shift @c;
        $R = shift @c;
        # ensure the left col is alphabetically less than the right col
        if $L.c gt $R.c  {
            ($L, $R) = ($R, $L);
        }
    }
    elsif @c.head.c eq @c.tail.c {
        $range-type = IsCol;
        $T = shift @c;
        $B = shift @c;
        # ensure the top row is less than the bottom row
        if $T.r > $B.r  {
            ($T, $B) = ($B, $T);
        }
    }
    else {
        die "FATAL: the two cells '{@c.head.A1}' and '{@c.tail.A1}' are not a linear range";
    }

    my ($start-A1, $end-A1);
    if $range-type ~~ IsCol {
        $start-A1 = $T.A1;
        $end-A1   = $B.A1;
    }
    elsif $range-type ~~ IsRow {
        $start-A1 = $L.A1;
        $end-A1   = $R.A1;
    }

    =begin comment
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
    =end comment

    my ($start-row, $start-col, $srow-abs, $scol-abs) = xl-cell-to-rowcol $start-A1;
    my ($end-row  , $end-col  , $erow-abs, $ecol-abs) = xl-cell-to-rowcol $end-A1;

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
        die "FATAL: unexpected non-linear cell range: '{$start-A1}:{$end-A1}'";
    }

    return @A1;

} # end of: sub split-linear

sub split-rectangular(@bcells, :$debug --> Array) is export {
    # the four @bcells are the bounding cells of a rectangular range
    die "FATAL: there should be four cells but there are {@bcells.elems}" if @bcells.elems != 4;

    # We assume the cells are such that they satisfy the following
    # rules:
    #   there must be two each of the two "A" (column) values
    #   there must be two each of the two "1" (row) values

    # The upper-left corner must have the "smallest" letter and number
    # and the lower-right must have the "largest."  The remaining
    # cells can be placed thusly: the cell with the same alpha as
    # upper-left must be the lower-left, and the remaining cell must
    # be the upper-right.

    # create an array of the cell objects:
    my @c;
    for @bcells -> $A1 {
        # if we have been rigorous in our plan the alpha chars
        # should be lower-case
        my $c = C.new: :$A1;
        @c.push: $c;
    }

    # sort the cells by row (numerically)
    my @a = @c.sort({$^b.r cmp $^a.r});

    my ($ul, $ur, $ll, $lr);
    # the top row
    $ul = shift @a;
    $ur = shift @a;
    # compare the columns alphabetically and swap if need be
    if $ul.c gt $ur.c {
        ($ul, $ur) = ($ur, $ul);
    }

    # the bottom row
    $ll = shift @a;
    $lr = shift @a;
    # compare the columns alphabetically and swap if need be
    if $ll.c gt $lr.c {
        ($ll, $lr) = ($lr, $ll);
    }

    my @A1;
    # step through each row and treat each as a linear range of
    # columns
    my $start-row = $ul.r;
    my $end-row   = $ll.r;
    for $start-row .. $end-row -> $rownum {
        # convert the row into a new linear range in "A1" notation
        my $start-cell = $ul.c ~ $rownum.Str;
        my $end-cell   = $ur.c ~ $rownum.Str;

        my @bcells = $start-cell, $end-cell;
        my @a1 = split-linear @bcells, :$debug;
        @A1.append: @a1; # flattens and adds the individual A1 names to the array
    }

    return @A1;

} # end of: sub split-rectangular

sub split-range(@cells, :$debug --> Array) is export {
    # Given a linear or rectangular range of "A1" cells,
    # split them into individual cells that make up the
    # entire range.

    if @cells.elems == 2 {
       return split-linear @cells, :$debug;
    }
    else {
       return split-rectangular @cells, :$debug;
    }

} # end of: sub split-range


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
