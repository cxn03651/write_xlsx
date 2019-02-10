# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar14 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar14
    @xlsx = 'chart_bar14.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    chart1     = workbook.add_chart(:type => 'bar', :embedded => 1)
    chart2     = workbook.add_chart(:type => 'bar', :embedded => 1)
    chart3     = workbook.add_chart(:type => 'column')

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids, [40294272, 40295808])
    chart2.instance_variable_set(:@axis_ids, [40261504, 65749760])
    chart3.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [65465728, 66388352])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet2.write('A1', data)
    worksheet2.write('A6', 'http://www.perl.com/')

    chart3.add_series(:values => '=Sheet2!$A$1:$A$5')
    chart3.add_series(:values => '=Sheet2!$B$1:$B$5')
    chart3.add_series(:values => '=Sheet2!$C$1:$C$5')

    chart1.add_series(:values => '=Sheet2!$A$1:$A$5')
    chart1.add_series(:values => '=Sheet2!$B$1:$B$5')
    chart1.add_series(:values => '=Sheet2!$C$1:$C$5')

    chart2.add_series(:values => '=Sheet2!$A$1:$A$5')

    worksheet2.insert_chart('E9',  chart1)
    worksheet2.insert_chart('F25', chart2)

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/charts/chart1.xml' => ['<c:pageMargins'],
        'xl/charts/chart2.xml' => ['<c:pageMargins'],
        'xl/charts/chart3.xml' => ['<c:pageMargins']
      }
    )
  end
end
