# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/table'

class TestTableWriteTableColumn < Test::Unit::TestCase
  def test_table_write_table_column
    expected = '<tableColumn id="1" name="Column1" />'

    table = Writexlsx::Package::Table.new
    table.__send__(:write_table_column, {:_name => 'Column1', :_id => 1})
    result = table.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
