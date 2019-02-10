# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_excel2003_style01
    @xlsx = 'excel2003_style01.xlsx'
    workbook    = WriteXLSX.new(@io, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    workbook.close
    compare_for_regression
  end
end
