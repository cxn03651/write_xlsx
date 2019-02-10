# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartDisplayUnits12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_display_units12
    @xlsx = 'chart_display_units12.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'scatter', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [93550464, 93548544])

    data = [
      [ 10000000, 20000000, 30000000, 20000000,  10000000 ]
    ]

    worksheet.write('A1', data)
    worksheet.write('B1', data)

    chart.add_series(
      :categories => '=Sheet1!$A$1:$A$5',
      :values     => '=Sheet1!$B$1:$B$5',
    )
    chart.set_y_axis(:display_units => 'hundreds', :display_units_visible => 0)
    chart.set_x_axis(:display_units => 'thousands', :display_units_visible => 0)

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
