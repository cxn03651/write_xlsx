# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestChartWritePageMargins < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_page_margins
    expected = '<c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>'
    result = @chart.__send__('write_page_margins')
    assert_equal(expected, result)
  end
end
