# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteFormatCode < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_format_code
    expected = '<c:formatCode>General</c:formatCode>'
    result = @chart.__send__('write_format_code', 'General')
    assert_equal(expected, result)
  end
end
