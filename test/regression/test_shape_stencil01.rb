# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShapeStencil01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shape_stencil01
    @xlsx = 'shape_stencil01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    worksheet.hide_gridlines(2)

    format      = workbook.add_format(:font => 'Arial', :size => 8)
    shape       = workbook.add_shape(
                                     :type   => 'rect',
                                     :width  => 90,
                                     :height => 90,
                                     :format => format
                                     )

    (1..10).each do |n|
      # Change the last 5 rectangles to stars.
      # Previously inserted shapes stay as rectangles
      shape.type = 'star5' if n == 6
      text = shape.type
      shape.text = [text, n.to_s].join(' ')
      worksheet.insert_shape('A1', shape, n * 100, 50)
    end

    workbook.close
    compare_for_regression
  end
end
