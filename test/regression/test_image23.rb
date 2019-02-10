# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage23 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image23
    @xlsx = 'image23.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('B2',
                           File.join(@test_dir, 'regression', 'images/black_72.jpg'))
    worksheet.insert_image('B8',
                           File.join(@test_dir, 'regression', 'images/black_96.jpg'))
    worksheet.insert_image('B13',
                           File.join(@test_dir, 'regression', 'images/black_150.jpg'))
    worksheet.insert_image('B17',
                           File.join(@test_dir, 'regression', 'images/black_300.jpg'))

    workbook.close
    compare_for_regression
  end
end
