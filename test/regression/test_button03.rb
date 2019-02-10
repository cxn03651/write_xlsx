# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button03
    @xlsx = 'button03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('C2', {})
    worksheet.insert_button('E5', {})

    workbook.close
    compare_for_regression
  end
end
