# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_image12
    @xlsx = 'image12.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.set_row(1, 75)
    worksheet.set_column('C:C', 32)

    worksheet.insert_image('C2',
                           'test/regression/images/logo.png')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
