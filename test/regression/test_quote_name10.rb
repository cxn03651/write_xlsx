# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionQuoteName10 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_quote_name10
    @xlsx = 'quote_name10.xlsx'
    workbook  = WriteXLSX.new(@io)

    # Test quoted/non-quoted sheet names.
    worksheet = workbook.add_worksheet("Sh.eet.1")
    chart = workbook.add_chart(type: 'column', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [46905600, 46796800])

    data = [
      [1, 2, 3,  4,  5],
      [2, 4, 6,  8, 10],
      [3, 6, 9, 12, 15]
    ]

    worksheet.write('A1', data)
    worksheet.repeat_rows(0, 1)
    worksheet.set_portrait
    worksheet.vertical_dpi = 200

    chart.add_series(values: ["Sh.eet.1", 0, 4, 0, 0])
    chart.add_series(values: ["Sh.eet.1", 0, 4, 1, 1])
    chart.add_series(values: ["Sh.eet.1", 0, 4, 2, 2])

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
