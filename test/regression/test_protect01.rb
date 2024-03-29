# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionProtect01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_protect01
    @xlsx = 'protect01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    unlocked = workbook.add_format(locked: 0, hidden: 0)
    hidden   = workbook.add_format(locked: 0, hidden: 1)

    worksheet.write('A1', 1)
    worksheet.write('A2', 2, unlocked)
    worksheet.write('A3', 3, hidden)

    workbook.close
    compare_for_regression
  end
end
