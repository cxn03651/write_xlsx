# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/table'

class TestTableWriteTableStyleInfo < Test::Unit::TestCase
  def test_table_write_table_style_info
    expected = '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'

    table = Writexlsx::Package::Table.new
    table.
      instance_variable_set(:@properties,
                            {
                              :_style            => 'TableStyleMedium9',
                              :_show_first_col   => 0,
                              :_show_last_col    => 0,
                              :_show_row_stripes => 1,
                              :_show_col_stripes => 0
                            }
                            )
    table.__send__(:write_table_style_info)
    result = table.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
