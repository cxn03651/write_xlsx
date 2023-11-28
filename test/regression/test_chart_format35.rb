# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartFormat35 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_format35
    @xlsx = 'chart_format35.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(type: 'line', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [52978048, 52980352])

    data = [
      [1, 2, 3, 4,  5],
      [2, 4, 6, 8,  10],
      [3, 6, 9, 12, 15]
    ]

    worksheet.write('A1', data)

    chart.add_series(
      categories: '=Sheet1!$A$1:$A$5',
      values:     '=Sheet1!$B$1:$B$5',
      trendline:  {
        type:             'linear',
        display_equation: 1,
        label:            {
          fill:   { color: 'red' },
          border: { color: 'yellow' },
          font:   {
            bold:   1,
            italic: 0,
            color:  '#0070C0'
          }
        }
      }
    )

    chart.add_series(
      categories: '=Sheet1!$A$1:$A$5',
      values:     '=Sheet1!$C$1:$C$5'
    )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
