# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartScatter15 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_scatter15
    @xlsx = 'chart_scatter15.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'scatter', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [58843520, 58845440])

    data = [
            [ 'X', 1,  3  ],
            [ 'Y', 10, 30 ]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :categories => '=Sheet1!$A$2:$A$3',
                     :values     => '=Sheet1!$B$2:$B$3'
                     )

    chart.set_x_axis(:name => '=Sheet1!$A$1',
                     :name_font => {:italic => 1, :baseline => -1})
    chart.set_y_axis(:name => '=Sheet1!$B$1')

    worksheet.insert_chart('E9', chart)

    workbook.close

    compare_for_regression
  end
end
