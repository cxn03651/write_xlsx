# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFirstsheet01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_firstsheet01
    @xlsx = 'firstsheet01.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet
    worksheet7 = workbook.add_worksheet
    worksheet8 = workbook.add_worksheet
    worksheet9 = workbook.add_worksheet
    worksheet10 = workbook.add_worksheet
    worksheet11 = workbook.add_worksheet
    worksheet12 = workbook.add_worksheet
    worksheet13 = workbook.add_worksheet
    worksheet14 = workbook.add_worksheet
    worksheet15 = workbook.add_worksheet
    worksheet16 = workbook.add_worksheet
    worksheet17 = workbook.add_worksheet
    worksheet18 = workbook.add_worksheet
    worksheet19 = workbook.add_worksheet
    worksheet20 = workbook.add_worksheet

    worksheet8.set_first_sheet
    worksheet20.activate

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                [],
                                {}
                                )
  end
end
