# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartDataLabels33 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_data_labels33
    @xlsx = 'chart_data_labels33.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [65546112, 70217728])

    data = [
      [1,  2,  3,  4,  5],
      [2,  4,  6,  8, 10],
      [3,  6,  9, 12, 15],
      [10, 20, 30, 40, 50]
    ]

    worksheet.write('A1', data)

    chart.add_series(
      :values      => '=Sheet1!$A$1:$A$5',
      :data_labels => {
        :value  => 1,
        :custom => [
          { :font => {
            :bold => 1, :italic => 1, :baseline => -1
          } }
        ]
      }
    )

    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
