# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/shared_strings'

class TestWriteSst < Test::Unit::TestCase
  def test_write_sst
    @obj = Writexlsx::Package::SharedStrings.new
    @obj.index('1')
    @obj.index('2')
    5.times { @obj.index('3') }
    @obj.__send__('write_sst')
    result = @obj.instance_variable_get(:@writer).string
    expected = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">'
    assert_equal(expected, result)
  end
end
