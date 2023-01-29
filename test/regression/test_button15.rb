# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionButton15 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button15
    @xlsx = 'button15.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_button(
      'C2', { :description => 'Some alternative text' }
    )

    workbook.close
    compare_for_regression
  end
end
