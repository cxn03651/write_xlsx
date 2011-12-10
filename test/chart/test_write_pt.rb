# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWritePt < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_pt
    expected = '<c:pt idx="0"><c:v>1</c:v></c:pt>'
    @chart.__send__('write_pt', 0, 1)
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
