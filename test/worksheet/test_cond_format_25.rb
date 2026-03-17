# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestCondFormat25 < Minitest::Test
  def test_conditional_formatting_cell_without_format
    xml = worksheet_xml_string do |_workbook, worksheet|
      worksheet.conditional_formatting(
        'A1',
        type:     'cell',
        format:   nil,
        criteria: '==',
        value:    'Test A2'
      )
    end

    expected = <<~XML
  <conditionalFormatting sqref="A1">
    <cfRule type="cellIs" priority="1" operator="equal">
      <formula>"Test A2"</formula>
    </cfRule>
  </conditionalFormatting>
XML

    assert_worksheet_xml_includes(xml, expected)
    refute_includes(xml, 'dxfId')
  end
end
