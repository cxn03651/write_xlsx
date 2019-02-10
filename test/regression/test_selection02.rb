# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSelection02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_selection02
    @xlsx = 'selection02.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet

    worksheet1.set_selection(3, 2)
    worksheet2.set_selection(3, 2, 6, 6)
    worksheet3.set_selection(6, 6, 3, 2)
    worksheet4.set_selection('C4')
    worksheet5.set_selection('C4:G7')
    worksheet6.set_selection('G7:C4')

    workbook.close
    compare_for_regression
  end
end
