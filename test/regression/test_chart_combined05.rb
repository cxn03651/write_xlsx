# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartCombined05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_combined05
    @xlsx = 'chart_combined05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart1    = workbook.add_chart(:type => 'bar',  :embedded => 1)
    chart2    = workbook.add_chart(:type => 'line', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids,  [60914304, 78899072])
    chart2.instance_variable_set(:@axis2_ids, [85542016, 85183872])

    data = [
      [ 2,  7,  3,  6,  2],
      [20, 25, 10, 10, 20]
    ]

    worksheet.write('A1', data)

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart2.add_series(:values => '=Sheet1!$B$1:$B$5', :y2_axis => 1)

    chart1.combine(chart2)

    worksheet.insert_chart('E9', chart1)

    workbook.close
    compare_for_regression(
      [],
      { 'xl/charts/chart1.xml' => [
          '<c:dispBlanksAs',
          '<c:tickLblPos',
          '<c:crosses',
          '<c:axPos'
        ] }
    )
  end
end
