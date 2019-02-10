# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartLine02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_line02
    @xlsx = 'chart_line02.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'line', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [63593856, 63612032])
    chart.instance_variable_set(:@axis2_ids, [63615360, 63613568])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 6, 8, 6, 4, 2 ]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5', :y2_axis => 1)

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/charts/chart1.xml' => ['<c:pageMargins'],
        'xl/workbook.xml'      => [ '<fileVersion' ]
      }
    )
  end
end
