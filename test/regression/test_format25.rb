# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat25 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format25
    @xlsx = 'format25.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1   = workbook.add_format(
      border_color: '#FF9966',
      border:       1
    )

    worksheet.write(2, 2, '', format1)

    workbook.close
    compare_for_regression
  end
end
