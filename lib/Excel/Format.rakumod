unit class Excel::Format;

use Excel::Utility;

# font and properties
has $.font      is rw;
has $.underline    is rw;
has $.italic    is rw;
has $.slanted    is rw;
has $.bold       is rw;

# justification
has $.l         is rw;
has $.c         is rw;
has $.r         is rw;

# cell properties
has $.color      is rw;
has $.background is rw; # color

# border properties
has $.linewidth is rw;
has $.top       is rw;
has $.bottom    is rw;
has $.left      is rw;
has $.right     is rw;

method write-format($perl-workbook, $row, $col, :$debug) is export {
}

#| The known formats we handle:
our %formats is export = set <
l
c
r
bold
linewidth
top
bottom
left
right
font
fontsize
>;

#| known colors
our %colors is export = set <
black
blue
brown
cyan
gray
green
lime
magenta
navy
orange
pink
purple
red
silver
white
yellow
>;
