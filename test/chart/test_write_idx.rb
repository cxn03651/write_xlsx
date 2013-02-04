# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteIdx < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_idx
    expected = '<c:idx val="0"/>'
    result = @chart.__send__('write_idx', 0)
    assert_equal(expected, result)
  end
end
