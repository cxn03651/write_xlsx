# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartLayout05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_layout05
    @xlsx = 'chart_layout05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'area', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [43495808, 43497728])

    data = [
            [1, 2, 3,  4,  5],
            [8, 7, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$B$1:$B$5'
                     )

    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$C$1:$C$5'
                     )

    chart.set_x_axis(
                     :name        => 'XXX',
                     :name_layout => {
                       :x => 0.34620319335083105,
                       :y => 0.85090259550889469
                     }
                     )

    chart.set_y_axis(
                     :name        => 'YYY',
                     :name_layout => {
                       :x => 0.21388888888888888,
                       :y => 0.26349919801691457
                     }
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
                                [],
                                { 'xl/charts/chart1.xml' => ['<c:pageMargins'] }
                                )
  end
end
