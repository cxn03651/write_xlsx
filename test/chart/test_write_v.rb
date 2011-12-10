# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteV < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_v
    expected = '<c:v>Apple</c:v>'
    @chart.__send__('write_v', 'Apple')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
