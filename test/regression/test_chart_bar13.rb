# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar13 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar13
    @xlsx = 'chart_bar13.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    chart1     = workbook.add_chart(:type => 'bar')
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    chart2     = workbook.add_chart(:type => 'bar')
    worksheet4 = workbook.add_worksheet

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [40294272, 40295808])
    chart2.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [62356096, 62366080])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet1.write('A1', data)

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart1.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart1.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart2.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart2.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart2.add_series(:values => '=Sheet1!$C$1:$C$5')

    workbook.close
    compare_for_regression
  end
end
