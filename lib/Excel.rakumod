unit module Excel;

# Perl modules
use File::LibMagic:from<Perl5>;
use Excel::Writer::XLSX:from<Perl5>;
use Spreadsheet::XLSX:from<Perl5>;

# older subs used with CVS::Parser 

sub read-xlsx($fnam, :$values) is export {
    # Returns an array of rows which are
    # arrays of columns of cells. Cell objects
    # are returned unless $values is true
    # in which case cell values are returned instead.
    my @rows;

    return @rows;
}

multi write-xlsx($fnam, @rows) is export {
    # Writes an xlsx file from an input 2x2 array
    # of xlsx cell objects.
}

multi write-xlsx($fnam, $workbook) is export {
    # Writes an xlsx file as a copy of the input 
    # Excel workbook.
}

sub write-xlsx-values($fnam, @rows) is export {
    # Writes an xlsx file from an input 2x2 array
    # of cells of text or numerical data.
}



