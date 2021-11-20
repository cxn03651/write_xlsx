# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionIgnoreError01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_ignore_error01
    @xlsx = 'ignore_error01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string('A1', '123')

    workbook.close
    compare_for_regression
  end
end
