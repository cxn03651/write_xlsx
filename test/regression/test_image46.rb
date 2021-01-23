# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage46 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image46
    @xlsx = 'image46.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/red.png'),
      0, 4
    )

    worksheet.set_row(8, 30, nil, 1)

    workbook.close
    compare_for_regression
  end
end
