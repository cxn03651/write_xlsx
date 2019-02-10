# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartPoints03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_points03
    @xlsx = 'chart_points03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'pie', :embedded => 1)

    workbook.set_custom_color(40, 0xCC, 0x00, 0x00)
    workbook.set_custom_color(41, 0x99, 0x00, 0x00)

    data = [
            [2, 5, 4]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :values => '=Sheet1!$A$1:$A$3',
                     :points => [
                                 { :fill => { :color => '#FF0000' } },
                                 { :fill => { :color => '#CC0000' } },
                                 { :fill => { :color => '#990000' } }
                                ]
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
