# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartDoughnut07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_doughnut07
    @xlsx = 'chart_doughnut07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart_doughnut = workbook.add_chart(type: 'doughnut', embedded: 1)

    worksheet.write_col('H2', ['Donut', 25, 50, 25, 100])
    worksheet.write_col('I2', ['Pie',   75,  1, 124])

    chart_doughnut.add_series(
      name:   '=Sheet1!$H$2',
      values: '=Sheet1!$H$3:$H$6'
    )

    chart_doughnut.add_series(
      name:   '=Sheet1!$I$2',
      values: '=Sheet1!$I$3:$I$6'
    )

    worksheet.insert_chart('E9', chart_doughnut)

    workbook.close
    compare_for_regression
  end
end
