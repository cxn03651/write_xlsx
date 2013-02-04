# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteMarkerSize < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_marker_size
    expected = '<c:size val="3"/>'
    result = @chart.__send__('write_marker_size', 3)
    assert_equal(expected, result)
  end
end
