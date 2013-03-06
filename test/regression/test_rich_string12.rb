# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRichString12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_rich_string12
    @xlsx = 'rich_string12.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.set_column('A:A', 30)
    worksheet.set_row(2, 60)

    bold   = workbook.add_format(:bold      => 1)
    italic = workbook.add_format(:italic    => 1)
    wrap   = workbook.add_format(:text_wrap => 1)

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('A2', 'Bar', italic)
    worksheet.write_rich_string('A3', "This is\n", bold, "bold\n", "and this is\n", italic, 'italic', wrap)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
