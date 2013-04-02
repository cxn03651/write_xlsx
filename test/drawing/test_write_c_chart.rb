# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteCChart < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_c_chart
    expected = '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'

    @drawing.__send__(:write_c_chart, 1)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
