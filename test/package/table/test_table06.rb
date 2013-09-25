# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable06 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table06
    # Set the table properties.
    @worksheet.add_table(
                         'C3:F13',
                         {
                           :columns => [
                                        {:header => 'Foo'},
                                        {:header => ''},
                                        {},
                                        {:header => 'Baz'}
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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
  <autoFilter ref="C3:F13"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Foo"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3"/>
    <tableColumn id="4" name="Baz"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
EOS
                      )
  end
end
