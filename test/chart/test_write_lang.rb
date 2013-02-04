# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteLang < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_lang
    expected = '<c:lang val="en-US"/>'
    result = @chart.__send__('write_lang')
    assert_equal(expected, result)
  end
end
