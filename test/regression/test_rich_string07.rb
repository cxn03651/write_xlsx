# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRichString07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_rich_string07
    @xlsx = 'rich_string07.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    bold   = workbook.add_format(:bold   => 1)
    italic = workbook.add_format(:italic => 1)

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('A2', 'Bar', italic)
    worksheet.write_rich_string('A3', 'a', bold, 'bc', 'defg')
    worksheet.write_rich_string('B4', 'abc', italic, 'de', 'fg')
    worksheet.write_rich_string('C5', 'a', bold, 'bc', 'defg')
    worksheet.write_rich_string('D6', 'abc', italic, 'de', 'fg')
    worksheet.write_rich_string('E7', 'a', bold, 'bcdef', 'g')
    worksheet.write_rich_string('F8', italic, 'abcd', 'efg')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
