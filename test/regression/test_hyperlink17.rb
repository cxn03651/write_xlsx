# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink17 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink17
    @xlsx = 'hyperlink17.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    # Test URL with whitespace.
    worksheet.write_url('A1', 'http://google.com/some link')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                {},
                                { 'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
