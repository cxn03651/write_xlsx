# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteOrder < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_order
    expected = '<c:order val="0"/>'
    result = @chart.__send__('write_order', 0)
    assert_equal(expected, result)
  end
end
