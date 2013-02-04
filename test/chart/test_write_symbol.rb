# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteSymbol < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_symbol
    expected = '<c:symbol val="none"/>'
    @chart.__send__('write_symbol', 'none')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
