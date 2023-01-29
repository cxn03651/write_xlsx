# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartColumn13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def test_chart_column13
    @xlsx = 'chart_column13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(type: 'column', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [60474496, 78612736])

    worksheet.write('A1', '1.1_1')
    worksheet.write('B1', '2.2_2')
    worksheet.write('A2', 1)
    worksheet.write('B2', 2)

    chart.add_series(
      categories: '=Sheet1!$A$1:$B$1',
      values:     '=Sheet1!$A$2:$B$2'
    )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
