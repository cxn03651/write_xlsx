# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartGap05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_gap05
    @xlsx = 'chart_gap05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [45938176, 59715584])
    chart.instance_variable_set(:@axis2_ids, [70848512, 54519680])

    data = [
            [1, 2, 3,  4,  5],
            [6, 8, 6,  4,  2]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :values     => '=Sheet1!$A$1:$A$5',
                     :gap        => 51,
                     :overlap    => 12
                     )

    chart.add_series(
                     :values     => '=Sheet1!$B$1:$B$5',
                     :y2_axis    => 1,
                     :gap        => 251,
                     :overlap    => -27
                     )

    chart.set_x2_axis(:label_position => 'next_to')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
                                [],
                                {
                                  'xl/charts/chart1.xml' => ['<c:pageMargins'],
                                  'xl/workbook.xml'      => [ '<fileVersion' ]
                                }
                                )
  end
end
