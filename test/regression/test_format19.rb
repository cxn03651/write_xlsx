# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat19 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format19
    @xlsx = 'format19.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(num_format: 'hh:mm;@')
    format2 = workbook.add_format(num_format: 'hh:mm;@', bg_color: 'yellow')

    worksheet.write(0, 0, 1, format1)
    worksheet.write(1, 0, 2, format2)

    workbook.close
    compare_for_regression
  end
end
