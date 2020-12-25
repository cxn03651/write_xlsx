# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteOrientation < Minitest::Test
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_orientation
    expected = '<c:orientation val="minMax"/>'
    result = @chart.__send__('write_orientation')
    assert_equal(expected, result)
  end
end
