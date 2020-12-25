# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage29 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image29
    @xlsx = 'image29.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      0, 10,
      File.join(@test_dir, 'regression', 'images/red_208.png'),
      -210, 1
    )

    workbook.close
    compare_for_regression
  end
end
