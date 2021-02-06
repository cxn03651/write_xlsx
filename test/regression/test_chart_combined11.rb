# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartCombined11 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_combined11
    @xlsx = 'chart_combined11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart_doughnut = workbook.add_chart(:type => 'doughnut', :embedded => 1)
    chart_pie      = workbook.add_chart(:type => 'pie',      :embedded => 1)

    worksheet.write_col('H2', ['Donut', 25, 50, 25, 100])
    worksheet.write_col('I2', ['Pie',   75,  1, 124])

    chart_doughnut.add_series(
      :name   => '=Sheet1!$H$2',
      :values => '=Sheet1!$H$3:$H$6',
      :points => [
        { :fill => { :color => '#FF0000' } },
        { :fill => { :color => '#FFC000' } },
        { :fill => { :color => '#00B050' } },
        { :fill => { :none  => 1 } }
      ]
    )

    chart_doughnut.set_rotation(270)
    chart_doughnut.set_legend(:none => 1)
    chart_doughnut.set_chartarea(
      :border => { :none => 1 },
      :fill   => { :none => 1 }
    )

    chart_pie.add_series(
      :name   => '=Sheet1!$I$2',
      :values => '=Sheet1!$I$3:$I$6',
      :points => [
        { :fill => { :none  => 1 } },
        { :fill => { :color => '#FF0000' } },
        { :fill => { :none  => 1 } }
      ]
    )

    chart_pie.set_rotation(270)

    chart_doughnut.combine(chart_pie)

    worksheet.insert_chart('A1', chart_doughnut)

    workbook.close
    compare_for_regression(
      [],
      { 'xl/charts/chart1.xml' => ['<c:dispBlanksAs'] }
    )
  end
end
