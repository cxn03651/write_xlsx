# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestGetChartRange < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_write_workbook_view_1
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>'
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_second_tab_selected
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" activeTab="1"/>'

    @workbook.activesheet = 1
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_second_tab_selected_first_sheet_set
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" firstSheet="2" activeTab="1"/>'

    @workbook.firstsheet  = 1
    @workbook.activesheet = 1
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_with_set_size
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>'

    @workbook.set_size
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_with_set_size_0_0
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>'

    @workbook.set_size(0, 0)
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_with_set_size_1073_644
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>'

    @workbook.set_size(1073, 644)
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_with_set_size_123_70
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="1845" windowHeight="1050"/>'

    @workbook.set_size(123, 70)
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end

  def test_write_workbook_view_with_set_size_719_490
    expected = '<workbookView xWindow="240" yWindow="15" windowWidth="10785" windowHeight="7350"/>'

    @workbook.set_size(719, 490)
    result   = @workbook.__send__('write_workbook_view')

    assert_equal(expected, result)
  end
end
