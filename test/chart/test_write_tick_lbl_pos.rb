# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteTickLabelPos < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_tick_label_pos
    expected = '<c:tickLblPos val="nextTo"/>'
    @chart.__send__('write_tick_label_pos', 'nextTo')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
