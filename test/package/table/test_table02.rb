# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable02 < Minitest::Test
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table02
    # Set the table properties.
    @worksheet.add_table('D4:I15', style: 'Table Style Light 17')
    @worksheet.__send__(:prepare_tables, 1, {})

    table = @worksheet.tables[0]
    table.__send__(:assemble_xml_file)

    result = got_to_array(table.instance_variable_get(:@writer).string)

    assert_equal(expected, result)
  end

  def expected
    expected_to_array(
      <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="D4:I15" totalsRowShown="0">
  <autoFilter ref="D4:I15"/>
  <tableColumns count="6">
    <tableColumn id="1" name="Column1"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3"/>
    <tableColumn id="4" name="Column4"/>
    <tableColumn id="5" name="Column5"/>
    <tableColumn id="6" name="Column6"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleLight17" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
EOS
    )
  end
end
