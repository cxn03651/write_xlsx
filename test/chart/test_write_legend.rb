# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteLegend < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_legend
    expected = '<c:legend><c:legendPos val="r"/><c:layout/></c:legend>'
    @chart.__send__('write_legend')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
