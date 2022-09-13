# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage30 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image30
    @xlsx = 'image30.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9',
      File.join(@test_dir, 'regression', 'images/red.png'),
      -2, -1
    )

    workbook.close
    compare_for_regression
  end
end
