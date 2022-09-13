# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionBackground03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background03
    @xlsx = 'background03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/logo.jpg')
    worksheet.set_background('test/regression/images/logo.jpg')

    workbook.close
    compare_for_regression
  end
end
