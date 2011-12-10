# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteNumFmt < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_num_fmt
    expected = '<c:numFmt formatCode="General" sourceLinked="1" />'
    @chart.instance_variable_set(:@has_category, 1)
    result = @chart.__send__('write_num_fmt')
    assert_equal(expected, result)
  end
end
