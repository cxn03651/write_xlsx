#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A simple example of how to use the WriteXLSX gem to
# demonstrate stenciling in an Excel xlsx file.
#
# reverse('c'), May 2012, John McNamara, jmcnamara@cpan.org
# converted to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('shape4.xlsx')
worksheet = workbook.add_worksheet

worksheet.hide_gridlines(2)

type = 'rect'
shape = workbook.add_shape(
  type:   type,
  width:  90,
  height: 90
)

(1..10).each do |n|
  # Change the last 5 rectangles to stars. Previously
  # inserted shapes stay as rectangles.
  type = 'star5' if n == 6
  shape.type = type
  shape.text = "#{type} #{n}"
  worksheet.insert_shape('A1', shape, n * 100, 50)
end

stencil = workbook.add_shape(
  stencil: 1,     # The default.
  width:   90,
  height:  90,
  text:    'started as a box'
)
worksheet.insert_shape('A1', stencil, 100, 150)

stencil.stencil = 0
worksheet.insert_shape('A1', stencil, 200, 150)
worksheet.insert_shape('A1', stencil, 300, 150)

# Ooopa! Changed my mind.
# Change the rectangle to an ellipse (circle),
# for the last two shapes.
stencil.type = 'ellipse'
stencil.text = 'Now its a circle'

workbook.close
