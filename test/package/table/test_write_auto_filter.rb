# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTableWriteAutoFilter01 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table_write_auto_filter
    expected = '<autoFilter ref="C3:F13"/>'

    table = Writexlsx::Package::Table.new(@worksheet, 1, 1, 2, 2)
    table.instance_variable_set(:@autofilter, 'C3:F13')

    table.__send__(:write_auto_filter)
    result = table.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
