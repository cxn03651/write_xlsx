# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable10 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table10
    # Set the table properties.
    @worksheet.add_table(
                         'C2:F13',
                         {:name => 'MyTable'}
                         )
    @worksheet.__send__(:prepare_tables, 1)

    table = @worksheet.tables[0]
    table.__send__(:assemble_xml_file)

    result = got_to_array(table.instance_variable_get(:@writer).string)

    assert_equal(expected, result)
  end

  def expected
    expected_to_array(
                      <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="MyTable" displayName="MyTable" ref="C2:F13" totalsRowShown="0">
  <autoFilter ref="C2:F13"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Column1"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3"/>
    <tableColumn id="4" name="Column4"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
EOS
                      )
  end
end
