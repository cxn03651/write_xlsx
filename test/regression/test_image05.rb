# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image05
    @xlsx = 'image05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('A1', 'test/regression/images/blue.png')
    worksheet.insert_image('B3', 'test/regression/images/red.jpg')
    worksheet.insert_image('D5', 'test/regression/images/yellow.jpg')
    worksheet.insert_image('F9', 'test/regression/images/grey.png')

    workbook.close
    compare_for_regression
  end
end
