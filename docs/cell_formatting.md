---
layout: default
title: Cell Formatting
---
### <a name="cell_formatting" class="anchor" href="#cell_formatting"><span class="octicon octicon-link" /></a>CELL FORMATTING

This section describes the methods and properties that are available for
formatting cells in Excel. The properties of a cell that can be formatted
include: fonts, colours, patterns, borders, alignment and number formatting.

##### <a name="creating_and_using_a_format_object" class="anchor" href="#creating_and_using_a_format_object"><span class="octicon octicon-link" /></a>Creating and using a Format object

Cell formatting is defined through a Format object.
Format objects are created by calling the
[`Workbook#add_format()`](workbook.html#add_format) method as follows:

    format1 = workbook.add_format              # Set properties later
    format2 = workbook.add_format(properties)   # Set at creation

The format object holds all the formatting properties that can be applied to a
cell, a row or a column. The process of setting these properties is discussed
in the next section.

Once a Format object has been constructed and its properties have been set
it can be passed as an argument to the worksheet write methods as follows:

    worksheet.write(0, 0, 'One', format)
    worksheet.write_string(1, 0, 'Two', format)
    worksheet.write_number(2, 0, 3, format)
    worksheet.write_blank(3, 0, format)

Formats can also be passed to the worksheet `set_row()` and `set_column()`
methods to define the default property for a row or column.

    worksheet.set_row(0, 15, format)
    worksheet.set_column(0, 0, 15, format)

##### <a name="format_methods_and_format_properties" class="anchor" href="#format_methods_and_format_properties"><span class="octicon octicon-link" /></a>Format methods and Format properties

The following table shows the Excel format categories, the formatting properties
that can be applied and the equivalent object method:

    Category   Description       Property        Method Name
    --------   -----------       --------        -----------
    Font       Font type         font            set_font()
               Font size         size            set_size()
               Font color        color           set_color()
               Bold              bold            set_bold()
               Italic            italic          set_italic()
               Underline         underline       set_underline()
               Strikeout         font_strikeout  set_font_strikeout()
               Super/Subscript   font_script     set_font_script()
               Outline           font_outline    set_font_outline()
               Shadow            font_shadow     set_font_shadow()

    Number     Numeric format    num_format      set_num_format()

    Protection Lock cells        locked          set_locked()
               Hide formulas     hidden          set_hidden()

    Alignment  Horizontal align  align           set_align()
               Vertical align    valign          set_align()
               Rotation          rotation        set_rotation()
               Text wrap         text_wrap       set_text_wrap()
               Justify last      text_justlast   set_text_justlast()
               Center across     center_across   set_center_across()
               Indentation       indent          set_indent()
               Shrink to fit     shrink          set_shrink()

    Pattern    Cell pattern      pattern         set_pattern()
               Background color  bg_color        set_bg_color()
               Foreground color  fg_color        set_fg_color()

    Border     Cell border       border          set_border()
               Bottom border     bottom          set_bottom()
               Top border        top             set_top()
               Left border       left            set_left()
               Right border      right           set_right()
               Border color      border_color    set_border_color()
               Bottom color      bottom_color    set_bottom_color()
               Top color         top_color       set_top_color()
               Left color        left_color      set_left_color()
               Right color       right_color     set_right_color()
               Diagonal type     diag_type       set_diag_type()
               Diagonal border   diag_border     set_diag_border()
               Diagonal color    diag_color      set_diag_color()

There are two ways of setting Format properties: by using the object method
interface or by setting the property directly. For example, a typical use of
the method interface would be as follows:

    format = workbook.add_format
    format.set_bold
    format.set_color('red')

By comparison the properties can be set directly by passing a hash of
properties to the Format constructor:

    format = workbook.add_format(bold: 1, color: 'red')

or after the Format has been constructed by means of the
`set_format_properties()` method as follows:

    format = workbook.add_format
    format.set_format_properties(bold: 1, color: 'red')

You can also store the properties in one or more named hashes and pass them
to the required method:

    font = {
        font:  'Calibri',
        size:  12,
        color: 'blue',
        bold:  1
    }

    shading = {
        bg_color: 'green',
        pattern:  1
    }


    format1 = workbook.add_format(font)            # Font only
    format2 = workbook.add_format(font, shading)   # Font and shading

The provision of two ways of setting properties might lead you to wonder which
is the best way. The method mechanism may be better if you prefer setting
properties via method calls (which the author did when the code was first
written) otherwise passing properties to the constructor has proved to be a
little more flexible and self documenting in practice. An additional advantage
of working with property hashes is that it allows you to share formatting
between workbook objects as shown in the example above.

##### <a name="working_with_formats" class="anchor" href="#working_with_formats"><span class="octicon octicon-link" /></a>Working with formats

The default format is Calibri 11 with all other properties off.

Each unique format in WriteXLSX must have a corresponding Format object.
It isn't possible to use a Format with a `write()` method and then redefine the
Format for use at a later stage. This is because a Format is applied to a cell
not in its current state but in its final state. Consider the following example:

    format = workbook.add_format
    format.set_bold
    format.set_color('red')
    worksheet.write('A1', 'Cell A1', format)
    format.set_color('green')
    worksheet.write('B1', 'Cell B1', format)

Cell A1 is assigned the Format `format` which is initially set to the colour
red. However, the colour is subsequently set to green.
When Excel displays Cell A1 it will display the final state of the Format
which in this case will be the colour green.

In general a method call without an argument will turn a property on, for example:

    format1 = workbook.add_format
    format1.set_bold       # Turns bold on
    format1.set_bold(1)    # Also turns bold on
    format1.set_bold(0)    # Turns bold off
