# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage36 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image36
    @xlsx = 'image36.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/happy.jpg')
    )

    workbook.close
    compare_for_regression
  end
end
