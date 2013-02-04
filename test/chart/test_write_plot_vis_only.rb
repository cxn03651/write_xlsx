# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWritePlotVisOnly < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_plot_vis_only
    expected = '<c:plotVisOnly val="1"/>'
    result = @chart.__send__('write_plot_vis_only')
    assert_equal(expected, result)
  end
end
