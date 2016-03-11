# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink21 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink21
    @xlsx = 'hyperlink21.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1', 'external:C:\Temp\Test 1')

    workbook.close

    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
