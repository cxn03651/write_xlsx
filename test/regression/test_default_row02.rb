# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_default_row02
    @xlsx = 'default_row02.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(nil, 1)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')

    (1..8).each { |row| worksheet.set_row(row, 15) }

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
