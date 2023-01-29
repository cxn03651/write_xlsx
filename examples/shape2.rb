#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A simple example of how to use the WriteXLSX gem to
# modify shapes properties in an Excel xlsx file.
#
# reverse('c'), May 2012, John McNamara, jmcnamara@cpan.org
# converted to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('shape2.xlsx')
worksheet = workbook.add_worksheet

worksheet.hide_gridlines(2)

plain = workbook.add_shape(
  type:   'smileyFace',
  text:   "Plain",
  width:  100,
  height: 100
)

bbformat = workbook.add_format(
  color: 'red',
  font:  'Lucida Calligraphy'
)

bbformat.set_bold
bbformat.set_underline
bbformat.set_italic

decor = workbook.add_shape(
  type:        'smileyFace',
  text:        'Decorated',
  rotation:    45,
  width:       200,
  height:      100,
  format:      bbformat,
  line_type:   'sysDot',
  line_weight: 3,
  fill:        'FFFF00',
  line:        '3366FF'
)

worksheet.insert_shape('A1', plain,  50, 50)
worksheet.insert_shape('A1', decor, 250, 50)

workbook.close
