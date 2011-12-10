# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteChartSpace < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_chart_space
    expected = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    result = @chart.__send__('write_chart_space')
    assert_equal(expected, result)
  end
end
