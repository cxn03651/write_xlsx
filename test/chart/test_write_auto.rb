# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteAuto < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_auto
    expected = '<c:auto val="1"/>'
    result = @chart.__send__('write_auto', 1)
    assert_equal(expected, result)
  end
end
