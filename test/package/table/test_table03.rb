# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable03 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table03
    # Set the table properties.
    @worksheet.add_table(
                         'C5:D16',
                         {
                           :banded_rows    => 0,
                           :first_column   => 1,
                           :last_column    => 1,
                           :banded_columns => 1
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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C5:D16" totalsRowShown="0">
  <autoFilter ref="C5:D16"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Column1"/>
    <tableColumn id="2" name="Column2"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="1" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>
</table>
EOS
                      )
  end
end
