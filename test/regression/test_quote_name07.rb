# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionQuoteName07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_quote_name07
    @xlsx = 'quote_name07.xlsx'
    workbook  = WriteXLSX.new(@io)

    # Test quoted/non-quoted sheet names.
    worksheet = workbook.add_worksheet("Sheet'1")
    chart = workbook.add_chart(type: 'column', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [48135552, 54701056])

    data = [
      [1, 2, 3,  4,  5],
      [2, 4, 6,  8, 10],
      [3, 6, 9, 12, 15]
    ]

    worksheet.write('A1', data)
    worksheet.repeat_rows(0, 1)
    worksheet.set_portrait
    worksheet.vertical_dpi = 200

    chart.add_series(values: ["Sheet'1", 0, 4, 0, 0])
    chart.add_series(values: ["Sheet'1", 0, 4, 1, 1])
    chart.add_series(values: ["Sheet'1", 0, 4, 2, 2])

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
