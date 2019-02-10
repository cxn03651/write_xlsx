# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartScatter08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_scatter08
    @xlsx = 'chart_scatter08.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(
                                     :type     => 'scatter',
                                     :embedded => 1
                                     )

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [103263232, 103261696])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A2', data)

    chart.add_series(
                     :categories => '=Sheet1!$A$2:$A$6',
                     :values     => '=Sheet1!$B$2:$B$6'
                     )

    chart.add_series(
                     :categories => '=Sheet1!$A$2:$A$6',
                     :values     => '=Sheet1!$C$2:$C$6'
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close

    compare_for_regression
                                # @xlsx,
                                # nil,
                                # {
                                #   'xl/charts/chart1.xml' => ['<c:pageMargins'],
                                #   'xl/workbook.xml'      => [ '<fileVersion', '<calcPr' ]
                                # }
                                # )

  end
end
