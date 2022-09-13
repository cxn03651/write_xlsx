# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage50 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image50
    @xlsx = 'image50.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9',  'test/regression/images/red.png')
    worksheet.insert_image('E13', 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
