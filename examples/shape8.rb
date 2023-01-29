#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A simple example of how to use the WriteXLSX gem to
# add shapes and one-to-many connectors to an Excel xlsx file.
#
# reverse('c'), May 2012, John McNamara, jmcnamara@cpan.org
# converted to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('shape8.xlsx')
worksheet = workbook.add_worksheet

# Add a circle, with centered text. c is for circle, not center.
cw = 60
ch = 60
cx = 210
cy = 190

ellipse = workbook.add_shape(
  type:   'ellipse',
  id:     2,
  text:   "Hello\nWorld",
  width:  cw,
  height: ch
)
worksheet.insert_shape('A1', ellipse, cx, cy)

# Add a plus sign at 4 different positionos around the circle.
pw = 20
ph = 20
px = 120
py = 250

plus = workbook.add_shape(
  type:   'plus',
  id:     3,
  width:  pw,
  height: ph
)

p1 = worksheet.insert_shape('A1', plus, 350, 150)
p2 = worksheet.insert_shape('A1', plus, 350, 350)
p3 = worksheet.insert_shape('A1', plus, 150, 350)
p4 = worksheet.insert_shape('A1', plus, 150, 150)

cxn_shape = workbook.add_shape(type: 'bentConnector3', fill: 0)

cxn_shape.start       = ellipse.id
cxn_shape.start_index = 2   # 2nd connection pt, clockwise from top(0).
cxn_shape.start_side  = 'r' # r)ight or b)ottom.

cxn_shape.end         = p1.id
cxn_shape.end_index   = 3   # 3rd connection point on plus, right side
cxn_shape.end_side    = 'l' # l)eft or t)op.
worksheet.insert_shape('A1', cxn_shape, 0, 0)

cxn_shape.end = p2.id
worksheet.insert_shape('A1', cxn_shape, 0, 0)

cxn_shape.end = p3.id
worksheet.insert_shape('A1', cxn_shape, 0, 0)

cxn_shape.end = p4.id
cxn_shape.adjustments = [-50, 45, 120]
worksheet.insert_shape('A1', cxn_shape, 0, 0)

workbook.close
