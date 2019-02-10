# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShapeConnect01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shape_connect01
    @xlsx = 'shape_connect01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:font => 'Arial', :size => 8)

    # Add a circle, with centered text
    ellipse = workbook.add_shape(:type => 'ellipse', :text => "Hello\nWorld",
                                 :width => 60, :height => 60, :format => format)
    worksheet.insert_shape('A1', ellipse, 50, 50)

    # Add a plus
    plus = workbook.add_shape(:type => 'plus', :width => 20, :height => 20)
    worksheet.insert_shape('A1', plus, 250, 200)

    # Create a bent connector to link the two shapes
    cxn_shape = workbook.add_shape(:type => 'bentConnector3')

    # Link the connector to the bottom of the circle
    cxn_shape.start = ellipse.id
    cxn_shape.start_index = 4    # 4th connection point, clockwise from top(0)
    cxn_shape.start_side  = 'b'  # r)ight or b)ottom

    # Link the connector to the bottom of the plus sign
    cxn_shape.end = plus.id
    cxn_shape.end_index = 0      # 0 - top connection point
    cxn_shape.end_side  = 't'    # l)eft of t)op

    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    workbook.close
    compare_for_regression
  end
end
