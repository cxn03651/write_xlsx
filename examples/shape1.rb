#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A simple example of how to use the WriteXLSX gem to
# add shapes to an Excel xlsx file.
#
# reverse('c'), May 2012, John McNamara, jmcnamara@cpan.org
# converted to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('shape1.xlsx')
worksheet = workbook.add_worksheet

# Add a circle, with centered text.
ellipse = workbook.add_shape(
  type:   'ellipse',
  text:   "Hello\nWorld",
  width:  60,
  height: 60
)
worksheet.insert_shape('A1', ellipse, 50, 50)

# Add a plus sign.
plus = workbook.add_shape(
  type:   'plus',
  width:  20,
  height: 20
)
worksheet.insert_shape('D8', plus)

workbook.close
