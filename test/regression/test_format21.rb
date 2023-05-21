# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat21 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format21
    @xlsx = 'format21.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1   = workbook.add_format(
      color:    'automatic',
      fg_color: 'automatic',
      bg_color: 'red',
      pattern:  6
    )

    worksheet.write(0, 0, 'Foo', format1)

    workbook.close
    compare_for_regression
  end
end
