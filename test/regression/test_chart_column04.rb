# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartColumn04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_column04
    @xlsx = 'chart_column04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [63591936, 63593856])
    chart.instance_variable_set(:@axis2_ids, [63613568, 63612032])

    data = [
            [1, 2, 3, 4, 5],
            [6, 8, 6, 4, 2]
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
        'xl/workbook.xml'      => ['<fileVersion']
      }
    )
  end
end
