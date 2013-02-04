# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWritePtCount < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_pt_count
    expected = '<c:ptCount val="5"/>'
    @chart.__send__('write_pt_count', 5)
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
