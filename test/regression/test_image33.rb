# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage33 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image33
    @xlsx = 'image33.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('D:D', 3.86)
    worksheet.set_column('E:E', 1.43)
    worksheet.set_row(7, 7.5)
    worksheet.set_row(8, 9.75)

    worksheet.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/red.png'),
      -2, -1
    )

    workbook.close
    compare_for_regression
  end
end
