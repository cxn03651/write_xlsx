# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat18 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format18
    @xlsx = 'format18.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    quote     = workbook.add_format(quote_prefix: 1)

    worksheet.write_string(0, 0, "= Hello", quote)

    workbook.close
    compare_for_regression
  end
end
