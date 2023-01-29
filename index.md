---
layout: default
title: WriteXLSX
---
### <a name="description" class="anchor" href="#description"><span class="octicon octicon-link" /></a>DESCRIPTION
The WriteXLSX rubygem can be used to create an Excel file in the 2007+ XLSX format.

The WriteXLSX and this document is ported from Perl module
[Excel::Wirter::XLSX](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/).
If you have any problem and question, please contact [me](mailto:nakamura.hideo@gmail.com).

Multiple worksheets can be added to a workbook and formatting can be applied to cells.
Text, numbers, and formulas can be written to the cells.
See [Examples](examples.html#examples)

WriteXLSX uses the same interface as the Writeexcel rubygem which produces
an Excel file in binary XLS format.

### <a name="requirements" class="anchor" href="#requirements"><span class="octicon octicon-link" /></a>REQUIREMENTS

WriteXLSX requires Ruby version 2.5.0 or later.

### <a name="synopsis" class="anchor" href="#synopsis"><span class="octicon octicon-link" /></a>SYNOPSIS

To write a string, a formatted string, a number and a formula to the first worksheet in an Excel workbook called ruby.xlsx:

    require 'write_xlsx'

    # Create a new Excel workbook
    workbook = WriteXLSX.new('ruby.xlsx')

    # Add a worksheet
    worksheet = workbook.add_worksheet

    #  Add and define a format
    format = workbook.add_format
    format.set_bold
    format.set_color('red')
    format.set_align('center')

    # Write a formatted and unformatted string, row and column notation.
    col = row = 0
    worksheet.write(row, col, 'Hi Excel!', format)
    worksheet.write(1, col, 'Hi Excel!')

    # Write a number and a formula using A1 notation
    worksheet.write('A3', 1.2345 )
    worksheet.write('A4', '=SIN(PI()/4)')

    # Write xlsx file to disk.
    workbook.close
