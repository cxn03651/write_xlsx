# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartLine07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_line07
    @xlsx = 'chart_line07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(
      type:     'line',
      embedded: 1
    )

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [77034624, 77036544])
    chart.instance_variable_set(:@axis2_ids,  [95388032, 103040896])

    data = [
      [1,  2,  3,  4,  5],
      [10, 40, 50, 20, 10],
      [1,  2,  3,  4,  5,  6,  7],
      [30, 10, 20, 40, 30, 10, 20]
    ]

    worksheet.write('A1', data)

    chart.add_series(
      categories: '=Sheet1!$A$1:$A$5',
      values:     '=Sheet1!$B$1:$B$5'
    )
    chart.add_series(
      categories: '=Sheet1!$C$1:$C$7',
      values:     '=Sheet1!$D$1:$D$7',
      y2_axis:    1
    )

    chart.set_x2_axis(label_position: 'next_to')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
      nil,
      'xl/charts/chart1.xml' => ['<c:crosses']
    )
  end
end
