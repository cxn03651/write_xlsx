# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'

class TestWriteCell < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_cell_0_0_n_1
    @worksheet.__send__('write_cell', 0, 0, [ 'n', 1 ])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<c r="A1"><v>1</v></c>'
    assert_equal(expected, result)
  end

  def test_write_cell_3_1_s_0
    @worksheet.__send__('write_cell', 3, 1, [ 's', 0 ])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<c r="B4" t="s"><v>0</v></c>'
    assert_equal(expected, result)
  end

  def test_write_cell_1_2_f_formula_nil_0
    format = nil
    @worksheet.__send__('write_cell', 1, 2, [ 'f', 'A3+A5', format, 0 ])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>'
    assert_equal(expected, result)
  end

  def test_write_cell_1_2_f_formula
    format = nil
    @worksheet.__send__('write_cell', 1, 2, [ 'f', 'A3+A5'])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>'
    assert_equal(expected, result)
  end

  def test_write_cell_0_0_a_formula_nil_a1_9500
    format = nil
    @worksheet.__send__('write_cell', 0, 0, [ 'a', 'SUM(B1:C1*B2:C2)', format, 'A1', 9500 ])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<c r="A1"><f t="array" ref="A1">SUM(B1:C1*B2:C2)</f><v>9500</v></c>'
    assert_equal(expected, result)
  end
end
