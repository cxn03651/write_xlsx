---
layout: default
title: Colours
---
### <a name="colors" class="anchor" href="#colors"><span class="octicon octicon-link" /></a>COLORS IN EXCEL

Excel provides a colour palette of 56 colours.
In WriteXLSX these colours are accessed via their palette index in the range 8..63.
This index is used to set the colour of fonts, cell patterns and cell borders.

For example:

    format = workbook.add_format(
               :color => 12, # index for blue
               :font  => 'Calibri',
               :size  => 12,
               :bold  => 1
             )

The most commonly used colours can also be accessed by name.
The name acts as a simple alias for the colour index:

    black     =>    8
    blue      =>   12
    brown     =>   16
    cyan      =>   15
    gray      =>   23
    green     =>   17
    lime      =>   11
    magenta   =>   14
    navy      =>   18
    orange    =>   53
    pink      =>   33
    purple    =>   20
    red       =>   10
    silver    =>   22
    white     =>    9
    yellow    =>   13

For example:

    font = workbook.add_format(:color => 'red')

Users of VBA in Excel should note that the equivalent colour indices are
in the range 1..56 instead of 8..63.

If the default palette does not provide a required colour you can override
one of the built-in values.
This is achieved by using the `set_custom_color()` workbook method to adjust
the RGB (red green blue) components of the colour:

    ferrari = workbook.set_custom_color(40, 216, 12, 12)

    format = workbook.add_format(
      :bg_color => ferrari,
      :pattern  => 1,
      :border   => 1
    )

    worksheet.write_blank('A1', format)

You can generate and example of the Excel palette using
`colors.rb` in the examples directory.

