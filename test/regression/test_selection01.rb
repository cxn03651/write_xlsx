# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSelection01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_selection01
    @xlsx = 'selection01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_selection('B4:C5')

    workbook.close
    compare_for_regression
  end
end
