# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink31 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink31
    @xlsx = 'hyperlink31.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format1   = workbook.add_format(:bold => 1)

    worksheet.write('A1', 'Test', format1)
    worksheet.write('A3', 'http://www.python.org/')

    workbook.close

    compare_for_regression
  end
end
