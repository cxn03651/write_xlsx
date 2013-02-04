# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWritePane < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_panes_freeze_1
    @worksheet.freeze_panes(1)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_0_1
    @worksheet.freeze_panes(0, 1)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_1_1
    @worksheet.freeze_panes(1, 1)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_1_0_19
    @worksheet.freeze_panes(1, 0, 19)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane ySplit="1" topLeftCell="A20" activePane="bottomLeft" state="frozen"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_G4
    @worksheet.freeze_panes('G4')
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_3_6_3_6_1
    @worksheet.freeze_panes(3, 6, 3, 6, 1)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozenSplit"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_15
    @worksheet.split_panes(15)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane ySplit="600" topLeftCell="A2"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_30
    @worksheet.split_panes(30)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane ySplit="900" topLeftCell="A3"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_105
    @worksheet.split_panes(105)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane ySplit="2400" topLeftCell="A8"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_0_843
    @worksheet.split_panes(0, 8.43)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="1350" topLeftCell="B1"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_0_1757
    @worksheet.split_panes(0, 17.57)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="2310" topLeftCell="C1"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_0_45
    @worksheet.split_panes(0, 45)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="5190" topLeftCell="F1"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_15_843
    @worksheet.split_panes(15, 8.43)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="1350" ySplit="600" topLeftCell="B2"/>'
    assert_equal(expected, result)
  end

  def test_write_panes_freeze_split_45_5114
    @worksheet.split_panes(45, 54.14)
    @worksheet.__send__('write_panes')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pane xSplit="6150" ySplit="1200" topLeftCell="G4"/>'
    assert_equal(expected, result)
  end
end
