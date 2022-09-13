# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionBackground01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background01
    @xlsx = 'background01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/logo.jpg')

    workbook.close
    compare_for_regression
  end
end
