# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProtect03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_protect03
    @xlsx = 'protect03.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    unlocked = workbook.add_format(:locked => 0, :hidden => 0)
    hidden   = workbook.add_format(:locked => 0, :hidden => 1)

    worksheet.protect('password')

    worksheet.write('A1', 1)
    worksheet.write('A2', 2, unlocked)
    worksheet.write('A3', 3, hidden)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
