# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink52 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink52
    @xlsx = 'hyperlink52.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url(0, 0, 'dynamicsnav://www.example.com')

    workbook.close

    compare_for_regression
  end
end
