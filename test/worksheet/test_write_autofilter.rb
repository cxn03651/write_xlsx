# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteAutofilter < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_write_auto_filter_with_no_filter
    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"/>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_East
    filter = 'x == East'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_East_or_x_North
    filter = 'x == East or  x == North'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/><filter val="North"/></filters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_East_and_x_North
    filter = 'x == East and  x == North'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters and="1"><customFilter val="East"/><customFilter val="North"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_ne_East
    filter = 'x != East'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="East"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_start_with_S
    filter = 'x == S*'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="S*"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_not_start_with_S
    filter = 'x != S*'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="S*"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_end_with_h
    filter = 'x == *h'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="*h"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_not_end_with_h
    filter = 'x != *h'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*h"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_contains_o
    filter = 'x =~ *o*'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="*o*"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_not_contains_r
    filter = 'x !~ *r*'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('A', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*r*"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_1000
    filter = 'x == 1000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><filters><filter val="1000"/></filters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_ne_2000
    filter = 'x != 2000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="notEqual" val="2000"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_gt_3000
    filter = 'x > 3000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="3000"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_ge_4000
    filter = 'x >= 4000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="greaterThanOrEqual" val="4000"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_lt_5000
    filter = 'x < 5000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="5000"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_x_le_6000
    filter = 'x <= 6000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="lessThanOrEqual" val="6000"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_1000_le_x_and_x_le_2000
    filter = 'x >= 1000 and x <= 2000'

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column('C', filter)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="1000"/><customFilter operator="lessThanOrEqual" val="2000"/></customFilters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_matches_East
    matches = ['East']

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column_list('A', matches)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_matches_East_and_North
    matches = ['East', 'North']

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column_list('A', matches)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/><filter val="North"/></filters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end

  def test_write_auto_filter_with_filter_matches_four_values
    matches = %w(February January July June)

    worksheet = @workbook.add_worksheet('Sheet1')
    worksheet.autofilter('A1:D51')
    worksheet.filter_column_list('D', matches)
    worksheet.__send__('write_auto_filter')
    result = worksheet.instance_variable_get(:@writer).string
    expected = '<autoFilter ref="A1:D51"><filterColumn colId="3"><filters><filter val="February"/><filter val="January"/><filter val="July"/><filter val="June"/></filters></filterColumn></autoFilter>'
    assert_equal(expected, result)
  end
end
