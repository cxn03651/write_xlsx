# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/chart'

class TestWriteLegendPos < Minitest::Test
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_legend_pos_default
    expected = '<c:legendPos val="r"/>'
    result = @chart.__send__('write_legend_pos', 'r')

    assert_equal(expected, result)
  end

  def test_write_legend_overlay_top_right
    expected = '<c:legend><c:legendPos val="tr"/><c:layout/><c:overlay val="1"/></c:legend>'
    @chart.set_legend(:position => 'overlay_top_right')
    @chart.__send__('write_legend')
    result = @chart.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
