# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteNumberFormat < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_number_format
    expected = '<c:numFmt formatCode="General" sourceLinked="1" />'
    result = @chart.__send__('write_number_format')
    assert_equal(expected, result)
  end
end
