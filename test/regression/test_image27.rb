# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage27 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image27
    @xlsx = 'image27.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('B2',
                           File.join(@test_dir, 'regression', 'images/mylogo.png'))

    workbook.close
    compare_for_regression
  end
end
