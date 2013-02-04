# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestAxPos < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_axis_pos
    expected = '<c:axPos val="l"/>'
    result = @chart.__send__('write_axis_pos', 'l')
    assert_equal(expected, result)
  end
end
