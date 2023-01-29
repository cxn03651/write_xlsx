#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# Example of how to add sparklines to an Excel::Writer::XLSX file.
#
# Sparklines are small charts that fit in a single cell and are
# used to show trends in data. This example shows the majority of
# options that can be applied to sparklines.
#
# reverse ('(c)'), November 2011, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook   = WriteXLSX.new('sparklines2.xlsx')
worksheet1 = workbook.add_worksheet
worksheet2 = workbook.add_worksheet
bold       = workbook.add_format(bold: 1)
row = 1

# Set the columns widths to make the output clearer.
worksheet1.set_column('A:A', 14)
worksheet1.set_column('B:B', 50)
worksheet1.zoom = 150

# Headings.
worksheet1.write('A1', 'Sparkline',   bold)
worksheet1.write('B1', 'Description', bold)

###############################################################################
#
str = 'A default "line" sparkline.'

worksheet1.add_sparkline(
  {
    location: 'A2',
    range:    'Sheet2!A1:J1'
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'A default "column" sparkline.'

worksheet1.add_sparkline(
  {
    location: 'A3',
    range:    'Sheet2!A2:J2',
    type:     'column'
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'A default "win/loss" sparkline.'

worksheet1.add_sparkline(
  {
    location: 'A4',
    range:    'Sheet2!A3:J3',
    type:     'win_loss'
  }
)

worksheet1.write(row, 1, str)
row += 2

###############################################################################
#
str = 'Line with markers.'

worksheet1.add_sparkline(
  {
    location: 'A6',
    range:    'Sheet2!A1:J1',
    markers:  1
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Line with high and low points.'

worksheet1.add_sparkline(
  {
    location:   'A7',
    range:      'Sheet2!A1:J1',
    high_point: 1,
    low_point:  1
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Line with first and last point markers.'

worksheet1.add_sparkline(
  {
    location:    'A8',
    range:       'Sheet2!A1:J1',
    first_point: 1,
    last_point:  1
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Line with negative point markers.'

worksheet1.add_sparkline(
  {
    location:        'A9',
    range:           'Sheet2!A1:J1',
    negative_points: 1
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Line with axis.'

worksheet1.add_sparkline(
  {
    location: 'A10',
    range:    'Sheet2!A1:J1',
    axis:     1
  }
)

worksheet1.write(row, 1, str)
row += 2

###############################################################################
#
str = 'Column with default style (1).'

worksheet1.add_sparkline(
  {
    location: 'A12',
    range:    'Sheet2!A2:J2',
    type:     'column'
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Column with style 2.'

worksheet1.add_sparkline(
  {
    location: 'A13',
    range:    'Sheet2!A2:J2',
    type:     'column',
    style:    2
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Column with style 3.'

worksheet1.add_sparkline(
  {
    location: 'A14',
    range:    'Sheet2!A2:J2',
    type:     'column',
    style:    3
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Column with style 4.'

worksheet1.add_sparkline(
  {
    location: 'A15',
    range:    'Sheet2!A2:J2',
    type:     'column',
    style:    4
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Column with style 5.'

worksheet1.add_sparkline(
  {
    location: 'A16',
    range:    'Sheet2!A2:J2',
    type:     'column',
    style:    5
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Column with style 6.'

worksheet1.add_sparkline(
  {
    location: 'A17',
    range:    'Sheet2!A2:J2',
    type:     'column',
    style:    6
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Column with a user defined colour.'

worksheet1.add_sparkline(
  {
    location:     'A18',
    range:        'Sheet2!A2:J2',
    type:         'column',
    series_color: '#E965E0'
  }
)

worksheet1.write(row, 1, str)
row += 2

###############################################################################
#
str = 'A win/loss sparkline.'

worksheet1.add_sparkline(
  {
    location: 'A20',
    range:    'Sheet2!A3:J3',
    type:     'win_loss'
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'A win/loss sparkline with negative points highlighted.'

worksheet1.add_sparkline(
  {
    location:        'A21',
    range:           'Sheet2!A3:J3',
    type:            'win_loss',
    negative_points: 1
  }
)

worksheet1.write(row, 1, str)
row += 2

###############################################################################
#
str = 'A left to right column (the default).'

worksheet1.add_sparkline(
  {
    location: 'A23',
    range:    'Sheet2!A4:J4',
    type:     'column',
    style:    20
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'A right to left column.'

worksheet1.add_sparkline(
  {
    location: 'A24',
    range:    'Sheet2!A4:J4',
    type:     'column',
    style:    20,
    reverse:  1
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
str = 'Sparkline and text in one cell.'

worksheet1.add_sparkline(
  {
    location: 'A25',
    range:    'Sheet2!A4:J4',
    type:     'column',
    style:    20
  }
)

worksheet1.write(row,   0, 'Growth')
worksheet1.write(row, 1, str)
row += 2

###############################################################################
#
str = 'A grouped sparkline. Changes are applied to all three.'

worksheet1.add_sparkline(
  {
    location: %w[A27 A28 A29],
    range:    ['Sheet2!A5:J5', 'Sheet2!A6:J6', 'Sheet2!A7:J7'],
    markers:  1
  }
)

worksheet1.write(row, 1, str)
row += 1

###############################################################################
#
# Create a second worksheet with data to plot.
#

worksheet2.set_column('A:J', 11)

data = [
  # Simple line data.
  [-2, 2, 3, -1, 0, -2, 3, 2, 1, 0],

  # Simple column data.
  [30, 20, 33, 20, 15, 5, 5, 15, 10, 15],

  # Simple win/loss data.
  [1, 1, -1, -1, 1, -1, 1, 1, 1, -1],

  # Unbalanced histogram.
  [5, 6, 7, 10, 15, 20, 30, 50, 70, 100],

  # Data for the grouped sparkline example.
  [-2, 2,  3, -1, 0, -2, 3, 2, 1, 0],
  [3,  -1, 0, -2, 3, 2,  1, 0, 2, 1],
  [0,  -2, 3, 2,  1, 0,  1, 2, 3, 1]
]

# Write the sample data to the worksheet.
worksheet2.write_col('A1', data)

workbook.close
