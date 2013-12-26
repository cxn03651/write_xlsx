# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteMajorGridlines < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_major_gridlines
    expected = '<c:majorGridlines/>'
    result = @chart.__send__('write_major_gridlines', Writexlsx::Chart::Gridline.new(:_visible => 1))
    assert_equal(expected, result)
  end
end
