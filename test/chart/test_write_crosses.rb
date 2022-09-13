# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/chart'

class TestWriteCrosses < Minitest::Test
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_crosses
    expected = '<c:crosses val="autoZero"/>'
    result = @chart.__send__('write_crosses', 'autoZero')
    assert_equal(expected, result)
  end
end
