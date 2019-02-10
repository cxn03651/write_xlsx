# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDate1904_01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_date_1904_01
    @xlsx = 'date_1904_01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    format      = workbook.add_format(:num_format => 14)

    worksheet.set_column('A:A', 12)

    worksheet.write_date_time('A1', '1900-01-00T', format)
    worksheet.write_date_time('A2', '1902-09-26T', format)
    worksheet.write_date_time('A3', '1913-09-08T', format)
    worksheet.write_date_time('A4', '1927-05-18T', format)
    worksheet.write_date_time('A5', '2173-10-14T', format)
    worksheet.write_date_time('A6', '4637-11-26T', format)

    workbook.close
    compare_for_regression
  end
end
