# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage57 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image57
    @xlsx = 'image57.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/logo.gif')

    workbook.close
    compare_for_regression
  end
end
