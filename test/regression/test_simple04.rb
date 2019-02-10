# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSimple04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_simple04
    @xlsx = 'simple04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column(0, 0, 12)

    format1 = workbook.add_format(:num_format => 20)
    format2 = workbook.add_format(:num_format => 14)

    worksheet.write_date_time(0, 0, 'T12:00', format1)
    worksheet.write_date_time(1, 0, '2013-01-27T', format2)

    workbook.close
    compare_for_regression
  end
end
