# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteMarker < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_marker
    expected = '<c:marker><c:symbol val="none"/></c:marker>'
    @chart.instance_variable_set(:@default_marker, Writexlsx::Chart::Marker.new(:type => 'none'))
    @chart.__send__('write_marker')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
