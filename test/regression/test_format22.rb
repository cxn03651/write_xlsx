# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat22 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format22
    @xlsx = 'format22.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1   = workbook.add_format(
      color:        'automatic',
      border:       1,
      border_color: 'automatic'
    )

    worksheet.write(0, 0, 'Foo', format1)

    workbook.close
    compare_for_regression
  end
end
