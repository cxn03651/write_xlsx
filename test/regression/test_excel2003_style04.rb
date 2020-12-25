# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_excel2003_style04
    @xlsx = 'excel2003_style04.xlsx'
    workbook    = WriteXLSX.new(@io, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.set_row(0, 21)

    workbook.close
    compare_for_regression
  end
end
