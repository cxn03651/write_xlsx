# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteLayout < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_idx
    expected = '<c:layout/>'
    result = @chart.__send__('write_layout')
    assert_equal(expected, result)
  end
end
