# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteCrossAx < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_cross_axis
    expected = '<c:crossAx val="82642816"/>'
    result = @chart.__send__('write_cross_axis', 82642816)
    assert_equal(expected, result)
  end
end
