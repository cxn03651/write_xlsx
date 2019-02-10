# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartPoints01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_points01
    @xlsx = 'chart_points01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'pie', :embedded => 1)

    data = [
            [2, 5, 4, 1, 7, 4]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :values => '=Sheet1!$A$1:$A$6',
                     :points => [
                                 { :fill => { :color => 'red' } }
                                ]
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
