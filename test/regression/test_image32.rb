# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage32 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image32
    @xlsx = 'image32.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Negative offset should be ignored.
    worksheet.insert_image(
      'B1',
      File.join(@test_dir, 'regression', 'images/red.png'),
      -100, -100
    )

    workbook.close
    compare_for_regression
  end
end
