# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionProtect06 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_protect06
    @xlsx = 'protect06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    unlocked = workbook.add_format(locked: 0, hidden: 0)
    hidden   = workbook.add_format(locked: 0, hidden: 1)

    worksheet.protect

    worksheet.unprotect_range('A1', nil, 'password')
    worksheet.unprotect_range('C1:C3')
    worksheet.unprotect_range('G4:I6', 'MyRange')
    worksheet.unprotect_range('K7', nil, 'foobar')

    worksheet.write('A1', 1)
    worksheet.write('A2', 2, unlocked)
    worksheet.write('A3', 3, hidden)

    workbook.close
    compare_for_regression
  end
end
