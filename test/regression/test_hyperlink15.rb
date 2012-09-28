# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink16 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink16
    @xlsx = 'hyperlink16.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write_url('B2', 'external:./subdir/blank.xlsx')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                {},
                                { 'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
