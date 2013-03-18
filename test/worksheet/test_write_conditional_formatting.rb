# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteConditionalFormatting < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_conditional_formatting_01
    format = Writexlsx::Format.new(Writexlsx::Formats.new)

    @worksheet.conditional_formatting('A1',
        :type     => 'cell',
        :format   => format,
        :criteria => 'greater than',
        :value    => 5
    )
    @worksheet.__send__('write_conditional_formats')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>'
    assert_equal(expected, result)
  end

  def test_conditional_formatting_02
    format = Writexlsx::Format.new(Writexlsx::Formats.new)

    @worksheet.conditional_formatting('A2',
        :type     => 'cell',
        :format   => format,
        :criteria => 'less than',
        :value    => 30
    )
    @worksheet.__send__('write_conditional_formats')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<conditionalFormatting sqref="A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>30</formula></cfRule></conditionalFormatting>'
    assert_equal(expected, result)
  end

  def test_conditional_formatting_03
    format = Writexlsx::Format.new(Writexlsx::Formats.new)

    @worksheet.conditional_formatting('A3',
        :type     => 'cell',
        :format   => nil,
        :criteria => '>=',
        :value    => 50
    )
    @worksheet.__send__('write_conditional_formats')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<conditionalFormatting sqref="A3"><cfRule type="cellIs" priority="1" operator="greaterThanOrEqual"><formula>50</formula></cfRule></conditionalFormatting>'
    assert_equal(expected, result)
  end

  def test_conditional_formatting_04
    format = Writexlsx::Format.new(Writexlsx::Formats.new)

    @worksheet.conditional_formatting('A1',
        :type     => 'cell',
        :format   => format,
        :criteria => 'between',
        :minimum  => 10,
        :maximum  => 20
    )
    @worksheet.__send__('write_conditional_formats')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="between"><formula>10</formula><formula>20</formula></cfRule></conditionalFormatting>'
    assert_equal(expected, result)
  end
end
