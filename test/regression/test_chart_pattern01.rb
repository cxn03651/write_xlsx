# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartPattern01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_pattern01
    @xlsx = 'chart_pattern01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [86421504, 86423040])

    data = [
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2]
    ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$3')
    chart.add_series(:values => '=Sheet1!$B$1:$B$3')
    chart.add_series(:values => '=Sheet1!$C$1:$C$3')
    chart.add_series(:values => '=Sheet1!$D$1:$D$3')
    chart.add_series(:values => '=Sheet1!$E$1:$E$3')
    chart.add_series(:values => '=Sheet1!$F$1:$F$3')
    chart.add_series(:values => '=Sheet1!$G$1:$G$3')
    chart.add_series(:values => '=Sheet1!$H$1:$H$3')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
