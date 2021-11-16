# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProtect05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_protect05
    @xlsx = 'protect05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    unlocked = workbook.add_format(:locked => 0, :hidden => 0)
    hidden   = workbook.add_format(:locked => 0, :hidden => 1)

    worksheet.protect

    worksheet.unprotect_range('A1')
    worksheet.unprotect_range('$C$1:$C$3')
    worksheet.unprotect_range('G4:I6', 'MyRange')
    worksheet.unprotect_range('K7')

    worksheet.write('A1', 1)
    worksheet.write('A2', 2, unlocked)
    worksheet.write('A3', 3, hidden)

    workbook.close
    compare_for_regression
  end
end
