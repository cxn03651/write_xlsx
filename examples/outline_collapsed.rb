#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#
# Example of how to use Excel::Writer::XLSX to generate Excel outlines and
# grouping.
#
# These examples focus mainly on collapsed outlines. See also the
# outlines.pl example program for more general examples.
#
# reverse ('(c)'), March 2008, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Create a new workbook and add some worksheets
workbook   = WriteXLSX.new('outline.xlsx')
worksheet1 = workbook.add_worksheet('Outlined Rows')
worksheet2 = workbook.add_worksheet('Collapsed Rows 1')
worksheet3 = workbook.add_worksheet('Collapsed Rows 2')
worksheet4 = workbook.add_worksheet('Collapsed Rows 3')
worksheet5 = workbook.add_worksheet('Outline Columns')
worksheet6 = workbook.add_worksheet('Collapsed Columns')

# Add a general format
bold = workbook.add_format(bold: 1)

###############################################################################
#
# Example 1: Create a worksheet with outlined rows. It also includes SUBTOTAL()
# functions so that it looks like the type of automatic outlines that are
# generated when you use the Excel Data->SubTotals menu item.
#

# The syntax is: set_row(row, height, XF, hidden, level, collapsed)
worksheet1.set_row(1, nil, nil, 0, 2)
worksheet1.set_row(2, nil, nil, 0, 2)
worksheet1.set_row(3, nil, nil, 0, 2)
worksheet1.set_row(4, nil, nil, 0, 2)
worksheet1.set_row(5, nil, nil, 0, 1)

worksheet1.set_row(6,  nil, nil, 0, 2)
worksheet1.set_row(7,  nil, nil, 0, 2)
worksheet1.set_row(8,  nil, nil, 0, 2)
worksheet1.set_row(9,  nil, nil, 0, 2)
worksheet1.set_row(10, nil, nil, 0, 1)

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet1, bold)

###############################################################################
#
# Example 2: Create a worksheet with collapsed outlined rows.
# This is the same as the example 1  except that the all rows are collapsed.
# Note: We need to indicate the row that contains the collapsed symbol '+' with
# the optional parameter, collapsed.

worksheet2.set_row(1, nil, nil, 1, 2)
worksheet2.set_row(2, nil, nil, 1, 2)
worksheet2.set_row(3, nil, nil, 1, 2)
worksheet2.set_row(4, nil, nil, 1, 2)
worksheet2.set_row(5, nil, nil, 1, 1)

worksheet2.set_row(6,  nil, nil, 1, 2)
worksheet2.set_row(7,  nil, nil, 1, 2)
worksheet2.set_row(8,  nil, nil, 1, 2)
worksheet2.set_row(9,  nil, nil, 1, 2)
worksheet2.set_row(10, nil, nil, 1, 1)

worksheet2.set_row(11, nil, nil, 0, 0, 1)

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet2, bold)

###############################################################################
#
# Example 3: Create a worksheet with collapsed outlined rows.
# Same as the example 1  except that the two sub-totals are collapsed.

worksheet3.set_row(1, nil, nil, 1, 2)
worksheet3.set_row(2, nil, nil, 1, 2)
worksheet3.set_row(3, nil, nil, 1, 2)
worksheet3.set_row(4, nil, nil, 1, 2)
worksheet3.set_row(5, nil, nil, 0, 1, 1)

worksheet3.set_row(6,  nil, nil, 1, 2)
worksheet3.set_row(7,  nil, nil, 1, 2)
worksheet3.set_row(8,  nil, nil, 1, 2)
worksheet3.set_row(9,  nil, nil, 1, 2)
worksheet3.set_row(10, nil, nil, 0, 1, 1)

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet3, bold)

###############################################################################
#
# Example 4: Create a worksheet with outlined rows.
# Same as the example 1  except that the two sub-totals are collapsed.

worksheet4.set_row(1, nil, nil, 1, 2)
worksheet4.set_row(2, nil, nil, 1, 2)
worksheet4.set_row(3, nil, nil, 1, 2)
worksheet4.set_row(4, nil, nil, 1, 2)
worksheet4.set_row(5, nil, nil, 1, 1, 1)

worksheet4.set_row(6,  nil, nil, 1, 2)
worksheet4.set_row(7,  nil, nil, 1, 2)
worksheet4.set_row(8,  nil, nil, 1, 2)
worksheet4.set_row(9,  nil, nil, 1, 2)
worksheet4.set_row(10, nil, nil, 1, 1, 1)

worksheet4.set_row(11, nil, nil, 0, 0, 1)

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet4, bold)

###############################################################################
#
# Example 5: Create a worksheet with outlined columns.
#
data = [
  %w[Month Jan Feb Mar Apr May Jun Total],
  ['North', 50,    20,    15,    25,    65,    80,   '=SUM(B2:G2)'],
  ['South', 10,    20,    30,    50,    50,    50,   '=SUM(B3:G3)'],
  ['East',  45,    75,    50,    15,    75,    100,  '=SUM(B4:G4)'],
  ['West',  15,    15,    55,    35,    20,    50,   '=SUM(B5:G6)']
]

# Add bold format to the first row
worksheet5.set_row(0, nil, bold)

# Syntax: set_column(col1, col2, width, XF, hidden, level, collapsed)
worksheet5.set_column('A:A', 10, bold)
worksheet5.set_column('B:G', 5, nil, 0, 1)
worksheet5.set_column('H:H', 10)

# Write the data and a formula
worksheet5.write_col('A1', data)
worksheet5.write('H6', '=SUM(H2:H5)', bold)

###############################################################################
#
# Example 6: Create a worksheet with collapsed outlined columns.
# This is the same as the previous example except collapsed columns.

# Add bold format to the first row
worksheet6.set_row(0, nil, bold)

# Syntax: set_column(col1, col2, width, XF, hidden, level, collapsed)
worksheet6.set_column('A:A', 10, bold)
worksheet6.set_column('B:G', 5,  nil, 1, 1)
worksheet6.set_column('H:H', 10, nil, 0, 0, 1)

# Write the data and a formula
worksheet6.write_col('A1', data)
worksheet6.write('H6', '=SUM(H2:H5)', bold)

workbook.close
