# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink21 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink21
    @xlsx = 'hyperlink21.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1', 'external:C:\Temp\Test 1')

    workbook.close

    compare_for_regression
  end
end
