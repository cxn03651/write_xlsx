# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable11 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
    @format = workbook.add_format
  end

  def test_table11
    # Set the table properties.
    @worksheet.add_table(
                         'C2:F14',
                         {
                           :total_row => 1,
                           :columns   => [
                                          {:total_string => 'Total'},
                                          {},
                                          {},
                                          {
                                            :total_function => 'count',
                                            :format         => @format,
                                            :formula        => 'SUM(Table1[[#This Row],[Column1]:[Column3]])'
                                          }
                                         ]
                         }
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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C2:F14" totalsRowCount="1">
  <autoFilter ref="C2:F13"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3"/>
    <tableColumn id="4" name="Column4" totalsRowFunction="count" dataDxfId="0">
      <calculatedColumnFormula>SUM(Table1[[#This Row],[Column1]:[Column3]])</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
EOS
                      )
  end
end
