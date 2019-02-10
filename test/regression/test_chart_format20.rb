# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartFormat20 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_format20
    @xlsx = 'chart_format20.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart1      = workbook.add_chart(:type => 'line', :embedded => 1)
    chart2      = workbook.add_chart(:type => 'line', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids, [80553856, 80555392])
    chart2.instance_variable_set(:@axis_ids, [84583936, 84585856])

    trend = {
      :type => 'linear',
      :line => {:color => 'red', :dash_type => 'dash'}
    }

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet.write('A1', data)

    chart1.add_series(
      :values    => '=Sheet1!$B$1:$B$5',
      :trendline => trend
    )
    chart1.add_series(:values => '=Sheet1!$C$1:$C$5')
    chart2.add_series(
      :values    => '=Sheet1!$B$1:$B$5',
      :trendline => trend
    )
    chart2.add_series(:values => '=Sheet1!$C$1:$C$5')

    worksheet.insert_chart('E9',  chart1)
    worksheet.insert_chart('E25', chart2)

    workbook.close
    compare_for_regression
  end
end
