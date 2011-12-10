#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'rubygems'
require 'write_xlsx'

workbook  = WriteXLSX.new('conditional_format.xlsx')
worksheet1 = workbook.add_worksheet

# Light red fill with dark red text.
format1 = workbook.add_format(
    :bg_color => '#FFC7CE',
    :color    => '#9C0006'
)

# Green fill with dark green text.
format2 = workbook.add_format(
    :bg_color => '#C6EFCE',
    :color    => '#006100'
)

# Some sample data to run the conditional formatting against.
data = [
    [ 90, 80,  50, 10,  20,  90,  40, 90,  30,  40 ],
    [ 20, 10,  90, 100, 30,  60,  70, 60,  50,  90 ],
    [ 10, 50,  60, 50,  20,  50,  80, 30,  40,  60 ],
    [ 10, 90,  20, 40,  10,  40,  50, 70,  90,  50 ],
    [ 70, 100, 10, 90,  10,  10,  20, 100, 100, 40 ],
    [ 20, 60,  10, 100, 30,  10,  20, 60,  100, 10 ],
    [ 10, 60,  10, 80,  100, 80,  30, 30,  70,  40 ],
    [ 30, 90,  60, 10,  10,  100, 40, 40,  30,  40 ],
    [ 80, 90,  10, 20,  20,  50,  80, 20,  60,  90 ],
    [ 60, 80,  30, 30,  10,  50,  80, 60,  50,  30 ]
]


# This example below highlights cells that have a value greater than or
# equal to 50 in red and cells below that value in green.

caption = 'Cells with values >= 50 are in light red. ' +
          'Values < 50 are in light green'

# Write the data.
worksheet1.write('A1', caption)
worksheet1.write_col('B3', data)

# Write a conditional format over a range.
worksheet1.conditional_formatting('B3:K12',
    {
        :type     => 'cell',
        :format   => format1,
        :criteria => '>=',
        :value    => 50
    }
)

# Write another conditional format over the same range.
worksheet1.conditional_formatting('B3:K12',
    {
        :type     => 'cell',
        :format   => format2,
        :criteria => '<',
        :value    => 50
    }
)

workbook.close
