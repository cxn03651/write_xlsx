# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink06
    @xlsx = 'hyperlink06.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1',  'external:C:\Temp\foo.xlsx')
    worksheet.write_url('A3',  'external:C:\Temp\foo.xlsx#Sheet1!A1')
    worksheet.write_url('A5',  'external:C:\Temp\foo.xlsx#Sheet1!A1', 'External', nil, 'Tip')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
