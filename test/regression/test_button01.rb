# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button01
    @xlsx = 'button01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('C2', {})

    workbook.close
    compare_for_regression
  end
end
