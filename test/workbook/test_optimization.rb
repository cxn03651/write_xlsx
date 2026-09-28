# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'

class TestOptimization < Minitest::Test
  def test_off_by_default
    workbook = WriteXLSX.new(StringIO.new)

    refute_predicate(workbook, :optimization?)
  end

  def test_constructor_option
    assert_predicate(WriteXLSX.new(StringIO.new, optimization: true), :optimization?)
    assert_predicate(WriteXLSX.new(StringIO.new, optimization: 1), :optimization?)
    refute_predicate(WriteXLSX.new(StringIO.new, optimization: false), :optimization?)
    refute_predicate(WriteXLSX.new(StringIO.new, optimization: 0), :optimization?)
  end

  def test_set_optimization
    workbook = WriteXLSX.new(StringIO.new)
    workbook.set_optimization

    assert_predicate(workbook, :optimization?)

    workbook.set_optimization(false)

    refute_predicate(workbook, :optimization?)
  end

  def test_set_optimization_after_add_worksheet
    workbook = WriteXLSX.new(StringIO.new)
    workbook.add_worksheet

    assert_raises(RuntimeError) { workbook.set_optimization }
  end

  def test_only_the_current_row_is_kept_in_memory
    workbook  = WriteXLSX.new(StringIO.new, optimization: true)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write('A2', 'Bar')
    worksheet.write('B2', 123)
    cell_data_store = worksheet.instance_variable_get(:@cell_data_store)

    assert_nil(cell_data_store[0])
    refute_nil(cell_data_store[1])
    assert_empty(workbook.shared_strings)

    workbook.close
  end

  def test_writes_to_already_written_rows_are_ignored
    workbook  = WriteXLSX.new(StringIO.new, optimization: true)
    worksheet = workbook.add_worksheet

    worksheet.write('A2', 'Foo')

    assert_equal(-2, worksheet.write_comment('A1', 'Some text'))
    assert_equal(-2, worksheet.set_row(0, 30))
    assert_equal(-2, worksheet.merge_range('A1:B1', 'Merged', workbook.add_format))

    workbook.close
  end

  def test_add_table_is_not_supported
    workbook  = WriteXLSX.new(StringIO.new, optimization: true)
    worksheet = workbook.add_worksheet

    _, err = capture_io do
      assert_equal(-1, worksheet.add_table('B3:F7'))
    end

    assert_match(/add_table\(\) isn't supported/, err)

    workbook.close
  end

  def test_chartsheet_is_not_optimized
    workbook   = WriteXLSX.new(StringIO.new, optimization: true)
    chartsheet = workbook.add_chart(type: 'column')

    refute_predicate(chartsheet, :optimization?)
    assert_nil(chartsheet.instance_variable_get(:@cell_data_file))

    workbook.close
  end

  def test_temporary_file_is_removed_after_close
    workbook  = WriteXLSX.new(StringIO.new, optimization: true)
    worksheet = workbook.add_worksheet
    worksheet.write('A1', 'Foo')
    path = worksheet.instance_variable_get(:@cell_data_file).path

    assert_path_exists(path)

    workbook.close

    refute_path_exists(path)
  end
end
