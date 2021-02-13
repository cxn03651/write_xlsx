# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage48 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image48
    @xlsx = 'image48.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/red.png')
    )

    worksheet2.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/red.png')
    )

    workbook.close
    compare_for_regression
  end
end
