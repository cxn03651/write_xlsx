# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartAxis06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_axis06
    @xlsx = 'chart_axis06.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'pie', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [47076480, 47078016])

    data = [
            [ 2,  4,  6],
            [60, 30, 10]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$3',
                     :values     => '=Sheet1!$B$1:$B$3'
                     )

    chart.set_title(:name => 'Title')
    # Axis formatting should be ignored.
    chart.set_x_axis(:name => 'XXX')
    chart.set_y_axis(:name => 'YYY')

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
