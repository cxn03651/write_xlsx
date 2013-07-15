---
layout: default
title: Page Set-Up Method
---

### <a name="page_set_up" class="anchor" href="#page_set_up"><span class="octicon octicon-link" /></a>PAGE SET-UP METHODS

Page set-up methods affect the way that a worksheet looks when it is printed.
They control features such as page headers and footers and margins.
These methods are really just standard worksheet methods.
They are documented here in a separate section for the sake of clarity.

The following methods are available for page set-up:

* [set_landscape](#set_landscape)
* [set_portrait](#set_portrait)
* [set_page_view](#set_page_view)
* [set_paper](#set_paper)
* [center_horizontally](#center_horizontally)
* [center_vertically](#center_vertically)
* [set_margins](#set_margins)
* [set_header](#set_header)
* [set_footer](#set_footer)
* [repeat_rows](#repeat_rows)
* [repeat_columns](#repeat_columns)
* [hide_gridlines](#hide_gridlines)
* [print_row_col_headers](#print_row_col_headers)
* [print_area](#print_area)
* [print_across](#print_across)
* [fit_to_pages](#fit_to_pages)
* [set_start_page](#set_start_page)
* [set_print_scale](#set_print_scale)
* [set_h_pagebreaks](#set_h_pagebreaks)
* [set_v_pagebreaks](#set_v_pagebreaks)

A common requirement when working with WriteXLSX is to apply the same
page set-up features to all of the worksheets in a workbook.
To do this you can use the `sheets()` method of the workbook class to access
the array of worksheets in a workbook:

    workbook.sheets.each do |worksheet|
      worksheet.set_landscape
    end

#### <a name="set_landscape" class="anchor" href="#set_landscape"><span class="octicon octicon-link" /></a>set_landscape()

This method is used to set the orientation of a worksheet's printed page to
landscape:

    worksheet.set_landscape    # Landscape mode

#### <a name="set_portrait" class="anchor" href="#set_portrait"><span class="octicon octicon-link" /></a>set_portrait()

This method is used to set the orientation of a worksheet's printed page to
portrait.
The default worksheet orientation is portrait, so you won't generally need
to call this method.

    worksheet.set_portrait    # Portrait mode

#### <a name="set_page_view" class="anchor" href="#set_page_view"><span class="octicon octicon-link" /></a>set_page_view()

This method is used to display the worksheet in "Page View/Layout" mode.

    worksheet.set_page_view

#### <a name="set_paper" class="anchor" href="#set_paper"><span class="octicon octicon-link" /></a>set_paper(index)

This method is used to set the paper format for the printed output of a worksheet.
The following paper styles are available:

    Index   Paper format            Paper size
    =====   ============            ==========
      0     Printer default         -
      1     Letter                  8 1/2 x 11 in
      2     Letter Small            8 1/2 x 11 in
      3     Tabloid                 11 x 17 in
      4     Ledger                  17 x 11 in
      5     Legal                   8 1/2 x 14 in
      6     Statement               5 1/2 x 8 1/2 in
      7     Executive               7 1/4 x 10 1/2 in
      8     A3                      297 x 420 mm
      9     A4                      210 x 297 mm
     10     A4 Small                210 x 297 mm
     11     A5                      148 x 210 mm
     12     B4                      250 x 354 mm
     13     B5                      182 x 257 mm
     14     Folio                   8 1/2 x 13 in
     15     Quarto                  215 x 275 mm
     16     -                       10x14 in
     17     -                       11x17 in
     18     Note                    8 1/2 x 11 in
     19     Envelope  9             3 7/8 x 8 7/8
     20     Envelope 10             4 1/8 x 9 1/2
     21     Envelope 11             4 1/2 x 10 3/8
     22     Envelope 12             4 3/4 x 11
     23     Envelope 14             5 x 11 1/2
     24     C size sheet            -
     25     D size sheet            -
     26     E size sheet            -
     27     Envelope DL             110 x 220 mm
     28     Envelope C3             324 x 458 mm
     29     Envelope C4             229 x 324 mm
     30     Envelope C5             162 x 229 mm
     31     Envelope C6             114 x 162 mm
     32     Envelope C65            114 x 229 mm
     33     Envelope B4             250 x 353 mm
     34     Envelope B5             176 x 250 mm
     35     Envelope B6             176 x 125 mm
     36     Envelope                110 x 230 mm
     37     Monarch                 3.875 x 7.5 in
     38     Envelope                3 5/8 x 6 1/2 in
     39     Fanfold                 14 7/8 x 11 in
     40     German Std Fanfold      8 1/2 x 12 in
     41     German Legal Fanfold    8 1/2 x 13 in

Note, it is likely that not all of these paper types will be available to the
end user since it will depend on the paper formats that the user's printer
supports. Therefore, it is best to stick to standard paper types.

    worksheet.set_paper(1)    # US Letter
    worksheet.set_paper(9)    # A4

If you do not specify a paper type the worksheet will print using the printer's
default paper.

#### <a name="center_horizontally" class="anchor" href="#center_horizontally"><span class="octicon octicon-link" /></a>center_horizontally()

Center the worksheet data horizontally between the margins on the printed page:

    worksheet.center_horizontally

#### <a name="center_vertically" class="anchor" href="#center_vertically"><span class="octicon octicon-link" /></a>center_vertically()

Center the worksheet data vertically between the margins on the printed page:

    worksheet.center_vertically

#### <a name="set_margins" class="anchor" href="#set_margins"><span class="octicon octicon-link" /></a>set_margins(inches)

There are several methods available for setting the worksheet margins on the
printed page:

    set_margins()        # Set all margins to the same value
    set_margins_LR()     # Set left and right margins to the same value
    set_margins_TB()     # Set top and bottom margins to the same value
    set_margin_left()    # Set left margin
    set_margin_right()   # Set right margin
    set_margin_top()     # Set top margin
    set_margin_bottom()  # Set bottom margin

All of these methods take a distance in inches as a parameter.
Note: 1 inch = 25.4mm. ;-)
The default left and right margin is 0.7 inch. The default top and bottom margin
is 0.75 inch.
Note, these defaults are different from the defaults used in the binary file
format by writeexcel gem.

#### <a name="set_header" class="anchor" href="#set_header"><span class="octicon octicon-link" /></a>set_header(string, margin)

Headers and footers are generated using a $string which is a combination of
plain text and control characters. The `margin` parameter is optional.

The available control character are:

    Control             Category            Description
    =======             ========            ===========
    &L                  Justification       Left
    &C                                      Center
    &R                                      Right

    &P                  Information         Page number
    &N                                      Total number of pages
    &D                                      Date
    &T                                      Time
    &F                                      File name
    &A                                      Worksheet name
    &Z                                      Workbook path

    &fontsize           Font                Font size
    &"font,style"                           Font name and style
    &U                                      Single underline
    &E                                      Double underline
    &S                                      Strikethrough
    &X                                      Superscript
    &Y                                      Subscript

    &&                  Miscellaneous       Literal ampersand &

Text in headers and footers can be justified (aligned) to the left,
center and right by prefixing the text with the control characters &L, &C and &R.

For example (with ASCII art representation of the results):

    worksheet.set_header('&LHello');

     ---------------------------------------------------------------
    |                                                               |
    | Hello                                                         |
    |                                                               |


    worksheet.set_header('&CHello');

     ---------------------------------------------------------------
    |                                                               |
    |                          Hello                                |
    |                                                               |


    worksheet.set_header('&RHello');

     ---------------------------------------------------------------
    |                                                               |
    |                                                         Hello |
    |                                                               |

For simple text, if you do not specify any justification the text will be
centred. However, you must prefix the text with &C if you specify a font name
or any other formatting:

    worksheet.set_header('Hello');

     ---------------------------------------------------------------
    |                                                               |
    |                          Hello                                |
    |                                                               |

You can have text in each of the justification regions:

    worksheet.set_header('&LCiao&CBello&RCielo');

     ---------------------------------------------------------------
    |                                                               |
    | Ciao                     Bello                          Cielo |
    |                                                               |

The information control characters act as variables that Excel will update as
the workbook or worksheet changes. Times and dates are in the users default format:

    worksheet.set_header('&CPage &P of &N');

     ---------------------------------------------------------------
    |                                                               |
    |                        Page 1 of 6                            |
    |                                                               |


    worksheet.set_header('&CUpdated at &T');

     ---------------------------------------------------------------
    |                                                               |
    |                    Updated at 12:30 PM                        |
    |                                                               |

You can specify the font size of a section of the text by prefixing it with the
control character &n where n is the font size:

    worksheet1.set_header('&C&30Hello Big')
    worksheet2.set_header('&C&10Hello Small')

You can specify the font of a section of the text by prefixing it with the
control sequence &"font,style" where fontname is a font name such as
"Courier New" or "Times New Roman" and style is one of the standard Windows
font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":

    worksheet1.set_header('&C&"Courier New,Italic"Hello')
    worksheet2.set_header('&C&"Courier New,Bold Italic"Hello')
    worksheet3.set_header('&C&"Times New Roman,Regular"Hello')

It is possible to combine all of these features together to create
sophisticated headers and footers. As an aid to setting up complicated headers
and footers you can record a page set-up as a macro in Excel and look at the
format strings that VBA produces.
Remember however that VBA uses two double quotes "" to indicate a single double
quote. For the last example above the equivalent VBA code looks like this:

    .LeftHeader   = ""
    .CenterHeader = "&""Times New Roman,Regular""Hello"
    .RightHeader  = ""

To include a single literal ampersand & in a header or footer you should use a
double ampersand &&:

    worksheet1.set_header('&CCuriouser && Curiouser - Attorneys at Law')

As stated above the margin parameter is optional. As with the other margins
the value should be in inches. The default header and footer margin is 0.3
inch. Note, the default margin is different from the default used in the binary
file format by Spreadsheet::WriteExcel. The header and footer margin size can
be set as follows:

    worksheet.set_header('&CHello', 0.75)

The header and footer margins are independent of the top and bottom margins.

Note, the header or footer string must be less than 255 characters.
Strings longer than this will not be written and a warning will be generated.

See, also the
[`headers.rb`](examples.html#headers)
program in the examples directory of the distribution.

#### <a name="set_footer" class="anchor" href="#set_footer"><span class="octicon octicon-link" /></a>set_footer(string, margin)

The syntax of the `set_footer()` method is the same as `set_header()`,
see above.

#### <a name="repeat_rows" class="anchor" href="#repeat_rows"><span class="octicon octicon-link" /></a>repeat_rows(first_row, last_row)

Set the number of rows to repeat at the top of each printed page.

For large Excel documents it is often desirable to have the first row or rows
of the worksheet print out at the top of each page.
This can be achieved by using the `repeat_rows()` method.
The parameters `first_row` and `last_row` are zero based.
The `last_row` parameter is optional if you only wish to specify one row:

    worksheet1.repeat_rows(0)    # Repeat the first row
    worksheet2.repeat_rows(0, 1) # Repeat the first two rows

#### <a name="repeat_columns" class="anchor" href="#repeat_columns"><span class="octicon octicon-link" /></a>repeat_columns(first_col, last_col)

Set the columns to repeat at the left hand side of each printed page.

For large Excel documents it is often desirable to have the first column or
columns of the worksheet print out at the left hand side of each page.
This can be achieved by using the `repeat_columns()` method.
The parameters `first_column` and `last_column` are zero based.
The `last_column` parameter is optional if you only wish to specify one column.
You can also specify the columns using A1 column notation,
see the note about ["Cell notation"](worksheet.html#cell-notation).

    worksheet1.repeat_columns(0)        # Repeat the first column
    worksheet2.repeat_columns(0, 1)     # Repeat the first two columns
    worksheet3.repeat_columns('A:A')    # Repeat the first column
    worksheet4.repeat_columns('A:B')    # Repeat the first two columns

#### <a name="hide_gridlines" class="anchor" href="#hide_gridlines"><span class="octicon octicon-link" /></a>hide_gridlines(option = 1)

This method is used to hide the gridlines on the screen and printed page.

Gridlines are the lines that divide the cells on a worksheet.
Screen and printed gridlines are turned on by default in an Excel worksheet.
If you have defined your own cell borders you may wish to hide the default gridlines.

    worksheet.hide_gridlines

The following values of $option are valid:

    0 : Don't hide gridlines
    1 : Hide printed gridlines only
    2 : Hide screen and printed gridlines

If you don't supply an argument or use nil, the default option is 1,
i.e. only the printed gridlines are hidden.

#### <a name="print_row_col_headers" class="anchor" href="#print_row_col_headers"><span class="octicon octicon-link" /></a>print_row_col_headers()

Set the option to print the row and column headers on the printed page.

An Excel worksheet looks something like the following;

     ------------------------------------------
    |   |   A   |   B   |   C   |   D   |  ...
     ------------------------------------------
    | 1 |       |       |       |       |  ...
    | 2 |       |       |       |       |  ...
    | 3 |       |       |       |       |  ...
    | 4 |       |       |       |       |  ...
    |...|  ...  |  ...  |  ...  |  ...  |  ...

The headers are the letters and numbers at the top and the left of the
worksheet. Since these headers serve mainly as a indication of position
on the worksheet they generally do not appear on the printed page.
If you wish to have them printed you can use the `print_row_col_headers()`
method :

    worksheet.print_row_col_headers

Do not confuse these headers with page headers as described in the
`set_header()` section above.

#### <a name="print_area" class="anchor" href="#print_area"><span class="octicon octicon-link" /></a>print_area(first_row, first_col, last_row, last_col)

This method is used to specify the area of the worksheet that will be printed.
All four parameters must be specified.
You can also use A1 notation, see the note about ["Cell notation"](worksheet.html#cell-notation).

    worksheet1.print_area('A1:H20')    # Cells A1 to H20
    worksheet2.print_area(0, 0, 19, 7) # The same
    worksheet2.print_area('A:H')       # Columns A to H if rows have data

#### <a name="print_across" class="anchor" href="#print_across"><span class="octicon octicon-link" /></a>print_across()

The print_across method is used to change the default print direction.
This is referred to by Excel as the sheet "page order".

    worksheet.print_across

The default page order is shown below for a worksheet that extends over 4 pages.
The order is called "down then across":

    [1] [3]
    [2] [4]

However, by using the print_across method the print order will be changed to
"across then down":

    [1] [2]
    [3] [4]

#### <a name="fit_to_pages" class="anchor" href="#fit_to_pages"><span class="octicon octicon-link" /></a>fit_to_pages(width, height)

The `fit_to_pages()` method is used to fit the printed area to a specific number
of pages both vertically and horizontally.
If the printed area exceeds the specified number of pages it will be scaled down
to fit. This guarantees that the printed area will always appear on the
specified number of pages even if the page size or margins change.

    worksheet1.fit_to_pages(1, 1)    # Fit to 1x1 pages
    worksheet2.fit_to_pages(2, 1)    # Fit to 2x1 pages
    worksheet3.fit_to_pages(1, 2)    # Fit to 1x2 pages

The print area can be defined using the `print_area()` method as described above.

A common requirement is to fit the printed output to n pages wide but have the
height be as long as necessary. To achieve this set the `height` to zero:

    worksheet1.fit_to_pages(1, 0)    # 1 page wide and as long as necessary

Note that although it is valid to use both `fit_to_pages()` and
`set_print_scale()` on the same worksheet only one of these options can be
active at a time. The last method call made will set the active option.

Note that `fit_to_pages()` will override any manual page breaks that are
defined in the worksheet.

Note: When using `fit_to_pages()` it may also be required to set the printer
paper size using `set_paper()` or else Excel will default to "US Letter".

#### <a name="set_start_page" class="anchor" href="#set_start_page"><span class="octicon octicon-link" /></a>set_start_page(start_page = 1)

The `set_start_page()` method is used to set the number of the starting page
when the worksheet is printed out. The default value is 1.

    worksheet.set_start_page(2)

#### <a name="set_print_scale" class="anchor" href="#set_print_scale"><span class="octicon octicon-link" /></a>set_print_scale(scale = 100)

Set the scale factor of the printed page.
Scale factors in the range 10 <= $scale <= 400 are valid:

    worksheet1.set_print_scale(50)
    worksheet2.set_print_scale(75)
    worksheet3.set_print_scale(300)
    worksheet4.set_print_scale(400)

The default scale factor is 100.

Note, `set_print_scale()` does not affect the scale of the visible page in Excel.
For that you should use `set_zoom()`.

Note also that although it is valid to use both `fit_to_pages()` and
`set_print_scale()` on the same worksheet only one of these options can be
active at a time. The last method call made will set the active option.

#### <a name="set_h_pagebreaks" class="anchor" href="#set_h_pagebreaks"><span class="octicon octicon-link" /></a>set_h_pagebreaks(breaks)

Add horizontal page breaks to a worksheet.

A page break causes all the data that follows it to be printed on the next page.
Horizontal page breaks act between rows.
To create a page break between rows 20 and 21 you must specify the break at
row 21. However in zero index notation this is actually row 20.
So you can pretend for a small while that you are using 1 index notation:

    worksheet1.set_h_pagebreaks(20)    # Break between row 20 and 21

The `set_h_pagebreaks()` method will accept a list of page breaks and you can
call it more than once:

    worksheet2.set_h_pagebreaks(20,  40,  60,  80,  100 )    # Add breaks
    worksheet2.set_h_pagebreaks(120, 140, 160, 180, 200 )    # Add some more

Note: If you specify the "fit to page" option via the `fit_to_pages()` method
it will override all manual page breaks.

There is a silent limitation of about 1000 horizontal page breaks per worksheet
in line with an Excel internal limitation.

#### <a name="set_v_pagebreaks" class="anchor" href="#set_v_pagebreaks"><span class="octicon octicon-link" /></a>set_v_pagebreaks(breaks)

Add vertical page breaks to a worksheet.

A page break causes all the data that follows it to be printed on the next
page. Vertical page breaks act between columns. To create a page break between
columns 20 and 21 you must specify the break at column 21. However in zero
index notation this is actually column 20. So you can pretend for a small while
that you are using 1 index notation:

    worksheet1.set_v_pagebreaks(20) # Break between column 20 and 21

The `set_v_pagebreaks()` method will accept a list of page breaks and you can
call it more than once:

    worksheet2.set_v_pagebreaks(20,  40,  60,  80,  100)    # Add breaks
    worksheet2.set_v_pagebreaks(120, 140, 160, 180, 200)    # Add some more

Note: If you specify the "fit to page" option via the `fit_to_pages()`
method it will override all manual page breaks.

