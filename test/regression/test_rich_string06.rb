# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRichString06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_rich_string06
    @xlsx = 'rich_string06.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    red = workbook.add_format(:color => 'red')

    worksheet.write('A1', 'Foo', red)
    worksheet.write('A2', 'Bar')
    worksheet.write_rich_string('A3', 'ab', red, 'cde', 'fg')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
