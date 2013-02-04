# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteLabelOffset < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_label_offset
    expected = '<c:lblOffset val="100"/>'
    result = @chart.__send__('write_label_offset', 100)
    assert_equal(expected, result)
  end
end
