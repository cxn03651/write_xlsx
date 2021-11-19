# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage52 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image52
    @xlsx = 'image52.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      :description => 'This is some alternative text'
    )

    workbook.close
    compare_for_regression
  end
end
