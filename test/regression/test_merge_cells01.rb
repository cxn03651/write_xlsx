# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionMergeCells01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_merge_cells01
    @xlsx = 'merge_cells01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    worksheet.set_selection('A4')

    worksheet.merge_range('A1:A2', 'col1', format)
    worksheet.merge_range('B1:B2', 'col2', format)
    worksheet.merge_range('C1:C2', 'col3', format)
    worksheet.merge_range('D1:D2', 'col4', format)

    workbook.close
    compare_for_regression
  end
end
