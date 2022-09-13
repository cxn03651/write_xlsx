# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteDefinedNames < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_write_defined_names_simple
    @workbook.instance_variable_set(:@defined_names, [['_xlnm.Print_Titles', 0, 'Sheet1!$1:$1']])
    @workbook.__send__('write_defined_names')
    result = @workbook.xml_str
    expected = '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName></definedNames>'
    assert_equal(expected, result)
  end

  def test_write_defined_names_multiple_range
    @workbook.add_worksheet
    @workbook.add_worksheet
    @workbook.add_worksheet('Sheet 3')

    @workbook.define_name("'Sheet 3'!Bar", "='Sheet 3'!$A$1")
    @workbook.define_name('Abc',           '=Sheet1!$A$1')
    @workbook.define_name('Baz',           '=0.98')
    @workbook.define_name('Sheet1!Bar',    '=Sheet1!$A$1')
    @workbook.define_name('Sheet2!Bar',    '=Sheet2!$A$1')
    @workbook.define_name('Sheet2!aaa',    '=Sheet2!$A$1')
    @workbook.define_name("'Sheet 3'!car", '="Saab 900"')
    @workbook.define_name('_Egg',          '=Sheet1!$A$1')
    @workbook.define_name('_Fog',          '=Sheet1!$A$1')

    @workbook.__send__('prepare_defined_names')
    @workbook.__send__('write_defined_names')

    result = got_to_array(@workbook.xml_str).join('')
    expected = %q(<definedNames><definedName name="_Egg">Sheet1!$A$1</definedName><definedName name="_Fog">Sheet1!$A$1</definedName><definedName name="aaa" localSheetId="1">Sheet2!$A$1</definedName><definedName name="Abc">Sheet1!$A$1</definedName><definedName name="Bar" localSheetId="2">'Sheet 3'!$A$1</definedName><definedName name="Bar" localSheetId="0">Sheet1!$A$1</definedName><definedName name="Bar" localSheetId="1">Sheet2!$A$1</definedName><definedName name="Baz">0.98</definedName><definedName name="car" localSheetId="2">"Saab 900"</definedName></definedNames>)
    assert_equal(expected, result)
  end
end
