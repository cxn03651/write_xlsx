# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage56 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image56
    @xlsx = 'image56.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/red.gif')

    workbook.close
    compare_for_regression
  end
end
