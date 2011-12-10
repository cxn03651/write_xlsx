# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteDefinedName < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_write_defined_name
    @workbook.__send__('write_defined_name', ['_xlnm.Print_Titles', 0, 'Sheet1!$1:$1'])
    result = @workbook.xml_str
    expected = '<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>'
    assert_equal(expected, result)
  end
end
