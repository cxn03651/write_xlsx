# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable09 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table09
    # Set the table properties.
    @worksheet.add_table(
                         'B2:K8',
                         {
                           :total_row => 1,
                           :columns => [
                                        {:total_string => 'Total'},
                                        {},
                                        {:total_function => 'Average'},
                                        {:total_function => 'COUNT'},
                                        {:total_function => 'count_nums'},
                                        {:total_function => 'max'},
                                        {:total_function => 'min'},
                                        {:total_function => 'sum'},
                                        {:total_function => 'std Dev'},
                                        {:total_function => 'var'}
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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:K8" totalsRowCount="1">
  <autoFilter ref="B2:K7"/>
  <tableColumns count="10">
    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3" totalsRowFunction="average"/>
    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
    <tableColumn id="5" name="Column5" totalsRowFunction="countNums"/>
    <tableColumn id="6" name="Column6" totalsRowFunction="max"/>
    <tableColumn id="7" name="Column7" totalsRowFunction="min"/>
    <tableColumn id="8" name="Column8" totalsRowFunction="sum"/>
    <tableColumn id="9" name="Column9" totalsRowFunction="stdDev"/>
    <tableColumn id="10" name="Column10" totalsRowFunction="var"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
EOS
                      )
  end
end
