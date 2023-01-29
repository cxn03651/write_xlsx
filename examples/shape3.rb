#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A simple example of how to use the WriteXLSX gem to
# scale shapes in an Excel xlsx file.
#
# reverse('c'), May 2012, John McNamara, jmcnamara@cpan.org
# converted to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('shape3.xlsx')
worksheet = workbook.add_worksheet

normal = workbook.add_shape(
  name:   'chip',
  type:   'diamond',
  text:   'Normal',
  width:  100,
  height: 100
)

worksheet.insert_shape('A1', normal, 50, 50)
normal.text = 'Scaled 3w x 2h'
normal.name = 'Hope'
worksheet.insert_shape('A1', normal, 250, 50, 3, 2)

workbook.close
