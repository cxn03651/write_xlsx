# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartFont01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_font01
    @xlsx = 'chart_font01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [43945344, 45705856])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart.set_title(:name => 'Title')

    chart.set_x_axis(
                     :name     => 'XXX',
                     :num_font => { :size => 11, :bold => 1, :italic => 1 }
                     )

    chart.set_y_axis(
                     :name     => 'YYY',
                     :num_font => { :size => 9, :bold => 0, :italic => 1 }
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
                 nil,
                 { 'xl/charts/chart1.xml' => ['<c:pageMargins'] }
                 )
  end
end
