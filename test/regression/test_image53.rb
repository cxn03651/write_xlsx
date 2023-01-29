# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage53 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image53
    @xlsx = 'image53.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      description: ''
    )

    workbook.close
    compare_for_regression
  end
end
