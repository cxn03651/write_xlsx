# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteExtLst < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_ext_lst
    @worksheet.__send__('write_ext_lst')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0" /></ext></extLst>'
    assert_equal(expected, result)
  end
end
