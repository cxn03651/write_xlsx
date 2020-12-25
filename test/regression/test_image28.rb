# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage28 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image28
    @xlsx = 'image28.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      0, 6,
      File.join(@test_dir, 'regression', 'images/red_208.png'),
      46, 1
    )

    workbook.close
    compare_for_regression
  end
end
