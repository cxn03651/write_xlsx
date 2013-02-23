# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_default_row03
    @xlsx = 'default_row03.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(24, 1)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')

    (1..8).each { |row| worksheet.set_row(row, 24) }

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
