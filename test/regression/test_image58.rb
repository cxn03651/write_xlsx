# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage58 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image58
    @xlsx = 'image58.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('A1048573', 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
