# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/table'

class TestTableWriteTableColumn < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table_write_table_column
    expected = '<tableColumn id="1" name="Column1"/>'

    table = Writexlsx::Package::Table.new(@worksheet, 1, 1, 2, 2)
    col_data = Writexlsx::Package::Table::ColumnData.new(1)

    table.__send__(:write_table_column, col_data)
    result = table.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
