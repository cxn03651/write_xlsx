# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage22 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image22
    @xlsx = 'image22.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('B2',
                           File.join(@test_dir, 'regression', 'images/black_300.jpg'))

    workbook.close
    compare_for_regression
  end
end
