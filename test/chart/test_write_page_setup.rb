# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestChartWritePageSetup < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_page_setup
    expected = '<c:pageSetup/>'
    result = @chart.__send__('write_page_setup')
    assert_equal(expected, result)
  end
end
