# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartDataLabels39 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_data_labels39
    @xlsx = 'chart_data_labels39.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(type: 'column', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [56179712, 56185600])

    data = [
      [1,  2,  3,  4,  5],
      [2,  4,  6,  8, 10],
      [3,  6,  9, 12, 15]
    ]

    worksheet.write('A1', data)

    chart.add_series(
      values:      '=Sheet1!$A$1:$A$5',
      data_labels: {
        value:    1,
        border:   {
          color:     'red',
          width:     1,
          dash_type: 'dash'
        },
        gradient: {
          colors: ['#DDEBCF', '#9CB86E', '#156B13']
        }
      }
    )

    chart.add_series(values: '=Sheet1!$B$1:$B$5')
    chart.add_series(values: '=Sheet1!$C$1:$C$5')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
