# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteLegendPos < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_legend_pos
    expected = '<c:legendPos val="r"/>'
    result = @chart.__send__('write_legend_pos', 'r')
    assert_equal(expected, result)
  end
end
