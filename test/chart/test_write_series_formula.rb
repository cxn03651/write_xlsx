# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteSeriesFormula < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_series_formula
    expected = '<c:f>Sheet1!$A$1:$A$5</c:f>'
    @chart.__send__('write_series_formula', 'Sheet1!$A$1:$A$5')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
