# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestRepeatFormula < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_repeat_formula
    format = nil

    expected = 'SUM(A1:A10)'
    row = 1
    col = 0
    formula = @worksheet.store_formula('=SUM(A1:A10)')
    @worksheet.repeat_formula(row, col, formula, format)
    result = @worksheet.instance_variable_get(:@cell_data_table)[row][col].token
    assert_equal(expected, result)

    expected = 'SUM(A2:A10)'
    row = 2
    col = 0
    formula = @worksheet.store_formula('=SUM(A1:A10)')
    @worksheet.repeat_formula(row, col, formula, format, 'A1', 'A2')
    result = @worksheet.instance_variable_get(:@cell_data_table)[row][col].token
    assert_equal(expected, result)

    expected = 'SUM(A2:A10)'
    row = 3
    col = 0
    formula = @worksheet.store_formula('=SUM(A1:A10)')
    @worksheet.repeat_formula(row, col, formula, format, /^A1$/, 'A2')
    result = @worksheet.instance_variable_get(:@cell_data_table)[row][col].token
    assert_equal(expected, result)

    expected = 'A2+A2'
    row = 4
    col = 0
    formula = @worksheet.store_formula('A1+A1')
    @worksheet.repeat_formula(row, col, formula, format, 'A1', 'A2', 'A1', 'A2')
    result = @worksheet.instance_variable_get(:@cell_data_table)[row][col].token
    assert_equal(expected, result)

    expected = 'A10 + SIN(A10)'
    row = 5
    col = 0
    formula = @worksheet.store_formula('A1 + SIN(A1)')
    @worksheet.repeat_formula(row, col, formula, format, /^A1$/, 'A10', /^A1$/, 'A10')
    result = @worksheet.instance_variable_get(:@cell_data_table)[row][col].token
    assert_equal(expected, result)
  end
end
