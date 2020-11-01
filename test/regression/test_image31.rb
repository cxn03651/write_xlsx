# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage31 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image31
    @xlsx = 'image31.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('D:D', 3.86)
    worksheet.set_row(7, 7.5)

    worksheet.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/red.png'),
      -2, -1
    )

    workbook.close
    compare_for_regression
  end
end
