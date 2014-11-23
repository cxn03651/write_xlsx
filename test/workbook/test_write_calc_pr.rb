# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteCalcPr < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_write_calc_pr
    @workbook.__send__('write_calc_pr')
    result = @workbook.xml_str
    expected = '<calcPr calcId="124519" fullCalcOnLoad="1"/>'
    assert_equal(expected, result)
  end

  def test_write_calc_pr_with_auto_except_tables_mode
    @workbook.set_calc_mode('auto_except_tables')
    @workbook.__send__('write_calc_pr')
    result = @workbook.xml_str
    expected = '<calcPr calcId="124519" calcMode="autoNoTable" fullCalcOnLoad="1"/>'
    assert_equal(expected, result)
  end

  def test_write_calc_pr_with_manual_mode
    @workbook.set_calc_mode('manual')
    @workbook.__send__('write_calc_pr')
    result = @workbook.xml_str
    expected = '<calcPr calcId="124519" calcMode="manual" calcOnSave="0"/>'
    assert_equal(expected, result)
  end

  def test_write_calc_pr_with_non_default_calc_id
    @workbook.set_calc_mode('auto', 12345)
    @workbook.__send__('write_calc_pr')
    result = @workbook.xml_str
    expected = '<calcPr calcId="12345" fullCalcOnLoad="1"/>'
    assert_equal(expected, result)
  end
end
