# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShapeConnect03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shape_connect03
    @xlsx = 'shape_connect03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Add a circle, with centered text. c is for circle, not center
    cw = 60
    ch = 60
    cx = 210
    cy = 190

    format = workbook.add_format(:font => 'Arial', :size => 8)
    ellipse = workbook.add_shape(
                                 :type => 'ellipse',
                                 :id => 2,
                                 :text => "Hello\nWorld",
                                 :width=> cw,
                                 :height => ch,
                                 :format => format
                                 )
    worksheet.insert_shape('A1', ellipse, cx, cy)

    # Add a plus sign at 4 different positions around the circle
    pw = 20
    ph = 20
    px = 120
    py = 250

    plus = workbook.add_shape(
                              :type => 'plus',
                              :id => 3,
                              :width => pw,
                              :height => ph
                              )
    p1 = worksheet.insert_shape('A1', plus, 350, 350)
    p2 = worksheet.insert_shape('A1', plus, 150, 350)
    p3 = worksheet.insert_shape('A1', plus, 350, 150)

    plus.adjustments = 35  # change shape of plus symbol
    p4 = worksheet.insert_shape('A1', plus, 150, 150)

    cxn_shape = workbook.add_shape(:type => 'bentConnector3', :fill=> 0)
    cxn_shape.start = ellipse.id
    cxn_shape.start_index = 4   # 4th connection point, clockwise from top(0)
    cxn_shape.start_side = 'b'  # r)ight or b)ottom

    cxn_shape.end = p1.id
    cxn_shape.end_index = 0     # first connection (zero based) from top
    cxn_shape.end_side = 't'    # l)eft or t)op

    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p2.id
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p3.id
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p4.id
    cxn_shape.adjustments = [-50, 45, 120]
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    workbook.close
    compare_for_regression
  end
end
