# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/table'

class TestTableWriteAutoFilter01 < Test::Unit::TestCase
  def test_table_write_auto_filter
    expected = '<autoFilter ref="C3:F13"/>'

    table = Writexlsx::Package::Table.new
    table.instance_variable_get(:@properties)[:_autofilter] = 'C3:F13'
    table.__send__(:write_auto_filter)
    result = table.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
