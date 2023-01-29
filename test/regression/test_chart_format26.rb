# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartFormat26 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_format26
    @xlsx = 'chart_format26.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(type: 'line', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [108652416, 108655744])

    data = [
      [1, 2, 3,  4,  5],
      [2, 4, 6,  8, 10],
      [3, 6, 9, 12, 15]
    ]

    worksheet.write('A1', data)

    chart.add_series(
      categories: '=Sheet1!$A$1:$A$5',
      values:     '=Sheet1!$B$1:$B$5',
      trendline:  { type: 'linear', display_equation: 1 }
    )

    chart.add_series(
      categories: '=Sheet1!$A$1:$A$5',
      values:     '=Sheet1!$C$1:$C$5'
    )

    chart.set_legend(delete_series: [2, 0])

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
