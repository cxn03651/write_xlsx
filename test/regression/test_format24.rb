# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat24 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format24
    @xlsx = 'format24.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1   = workbook.add_format(
      rotation: 270,
      indent:   1,
      align:    "center",
      valign:   "top"
    )

    worksheet.set_row(0, 75)

    worksheet.write(0, 0, 'ABCD', format1)

    workbook.close
    compare_for_regression
  end
end
