# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar15 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar15
    @xlsx = 'chart_bar15.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    chart1     = workbook.add_chart(:type => 'bar')
    worksheet2 = workbook.add_worksheet
    chart2     = workbook.add_chart(:type => 'column')

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [62576896, 62582784])
    chart2.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [65979904, 65981440])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet1.write('A1', data)
    worksheet2.write('A1', data)

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart1.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart1.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart2.add_series(:values => '=Sheet2!$A$1:$A$5')

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/charts/chart1.xml' => ['<c:pageMargins'],
        'xl/charts/chart2.xml' => ['<c:pageMargins']
      }
    )
  end
end
