# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage49 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image49
    @xlsx = 'image49.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.insert_image('A1', 'test/regression/images/blue.png')
    worksheet1.insert_image('B3', 'test/regression/images/red.jpg')
    worksheet1.insert_image('D5', 'test/regression/images/yellow.jpg')
    worksheet1.insert_image('F9', 'test/regression/images/grey.png')

    worksheet2.insert_image('A1', 'test/regression/images/blue.png')
    worksheet2.insert_image('B3', 'test/regression/images/red.jpg')
    worksheet2.insert_image('D5', 'test/regression/images/yellow.jpg')
    worksheet2.insert_image('F9', 'test/regression/images/grey.png')

    worksheet3.insert_image('A1', 'test/regression/images/blue.png')
    worksheet3.insert_image('B3', 'test/regression/images/red.jpg')
    worksheet3.insert_image('D5', 'test/regression/images/yellow.jpg')
    worksheet3.insert_image('F9', 'test/regression/images/grey.png')

    workbook.close
    compare_for_regression
  end
end
