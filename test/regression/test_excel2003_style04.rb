# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_excel2003_style04
    @xlsx = 'excel2003_style04.xlsx'
    workbook    = WriteXLSX.new(@xlsx, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.set_row(0, 21)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
