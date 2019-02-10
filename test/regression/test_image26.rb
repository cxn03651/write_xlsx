# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage26 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image26
    @xlsx = 'image26.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('B2',
                           File.join(@test_dir, 'regression', 'images/black_72.png'))
    worksheet.insert_image('B8',
                           File.join(@test_dir, 'regression', 'images/black_96.png'))
    worksheet.insert_image('B13',
                           File.join(@test_dir, 'regression', 'images/black_150.png'))
    worksheet.insert_image('B17',
                           File.join(@test_dir, 'regression', 'images/black_300.png'))

    workbook.close
    compare_for_regression
  end
end
