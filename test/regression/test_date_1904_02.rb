# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDate1904_02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_date_1904_02
    @xlsx = 'date_1904_02.xlsx'
    workbook    = WriteXLSX.new(@io)

    workbook.set_1904

    worksheet   = workbook.add_worksheet
    format      = workbook.add_format(:num_format => 14)

    worksheet.set_column('A:A', 12)

    worksheet.write_date_time('A1', '1904-01-01T', format)
    worksheet.write_date_time('A2', '1906-09-27T', format)
    worksheet.write_date_time('A3', '1917-09-09T', format)
    worksheet.write_date_time('A4', '1931-05-19T', format)
    worksheet.write_date_time('A5', '2177-10-15T', format)
    worksheet.write_date_time('A6', '4641-11-27T', format)

    workbook.close
    compare_for_regression
  end
end
