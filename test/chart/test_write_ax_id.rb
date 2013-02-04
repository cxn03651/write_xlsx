# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteAxId < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_axis_id
    expected = '<c:axId val="53850880"/>'
    result = @chart.__send__('write_axis_id', 53850880)
    assert_equal(expected, result)
  end
end
