# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'
require 'stringio'

#
# Test that Worksheet#write does not misdetect a multi-line string as a
# formula. In Ruby, /^=/ matches at the beginning of every line (unlike
# Perl, where the original Excel::Writer::XLSX code came from), so a
# multi-line string containing a line that starts with "=" used to be
# written as a formula cell, which corrupts the resulting xlsx file.
#
class TestWriteMultilineStringNotFormula < Minitest::Test
  def test_write_multiline_string_with_line_starting_with_equal_sign
    xml = worksheet_xml_string do |_workbook, worksheet|
      worksheet.write(0, 0, "first line\n=second line\nthird line")
    end

    refute_includes(xml, '<f>')
    assert_includes(xml, '<c r="A1" t="s">')
  end

  def test_write_multiline_string_with_array_formula_like_line
    xml = worksheet_xml_string do |_workbook, worksheet|
      worksheet.write(0, 0, "{=SUM(A1:A2)}\nsecond line")
    end

    refute_includes(xml, '<f')
    assert_includes(xml, '<c r="A1" t="s">')
  end

  def test_write_string_starting_with_equal_sign_is_still_a_formula
    xml = worksheet_xml_string do |_workbook, worksheet|
      worksheet.write(0, 0, '=1+1')
    end

    assert_includes(xml, '<f>1+1</f>')
  end

  def test_write_array_formula_string_is_still_an_array_formula
    xml = worksheet_xml_string do |_workbook, worksheet|
      worksheet.write(0, 0, '{=SUM(B1:C1*B2:C2)}')
    end

    assert_includes(xml, 't="array"')
    assert_includes(xml, 'SUM(B1:C1*B2:C2)')
  end
end
