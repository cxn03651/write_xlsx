# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage35 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image35
    @xlsx = 'image35.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/zero_dpi.jpg')
    )

    workbook.close
    compare_for_regression
  end
end
