# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionChartDisplayUnits11 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_display_units11
    @xlsx = 'chart_display_units11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(type: 'column', embedded: 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [69559424, 69560960])

    data = [
      [10000000, 20000000, 30000000, 20000000,  10000000]
    ]

    worksheet.write('A1', data)

    chart.add_series(values: '=Sheet1!$A$1:$A$5')
    chart.set_y_axis(display_units: 'hundreds', display_units_visible: 1)

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
