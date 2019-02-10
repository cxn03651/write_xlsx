# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar19 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar19
    @xlsx = 'chart_bar19.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet
    chart      = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [66558592, 66569344])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => ['Sheet1', 0, 4, 0, 0])
    chart.add_series(:values => ['Sheet1', 0, 4, 1, 1])
    chart.add_series(:values => ['Sheet1', 0, 4, 2, 2])

    chart.set_x_axis(:name => '=Sheet1!$A$2')
    chart.set_y_axis(:name => '=Sheet1!$A$3')
    chart.set_title(:name => '=Sheet1!$A$1')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/charts/chart1.xml' => ['<c:pageMargins']
      }
    )
  end
end
