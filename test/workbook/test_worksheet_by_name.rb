# -*- coding: utf-8 -*-

require 'helper'

class TestWorksheeByName < Minitest::Test
  def test_worksheet_by_name
    io = StringIO.new
    workbook = WriteXLSX.new(io)

    # Test a valid explicit name.
    expected = workbook.add_worksheet
    result   = workbook.worksheet_by_name('Sheet1')
    assert_equal(expected, result)
    result   = workbook.get_worksheet_by_name('Sheet1')
    assert_equal(expected, result)

    # Test a valid explicit name.
    expected = workbook.add_worksheet('Sheet 2')
    result   = workbook.worksheet_by_name('Sheet 2')
    assert_equal(expected, result)
    result   = workbook.get_worksheet_by_name('Sheet 2')
    assert_equal(expected, result)

    # Test an invalid name.
    result   = workbook.worksheet_by_name('Sheet3')
    assert_nil(result)
    result   = workbook.get_worksheet_by_name('Sheet3')
    assert_nil(result)

    # Test an invalid name.
    result   = workbook.worksheet_by_name
    assert_nil(result)
    result   = workbook.get_worksheet_by_name
    assert_nil(result)
  end
end
