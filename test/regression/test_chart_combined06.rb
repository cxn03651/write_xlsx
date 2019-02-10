# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartCombined06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_combined06
    @xlsx = 'chart_combined06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart1    = workbook.add_chart(:type => 'area',   :embedded => 1)
    chart2    = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids, [91755648, 91757952])
    chart2.instance_variable_set(:@axis_ids, [91755648, 91757952])

    data = [
      [ 2,  7,  3,  6,  2],
      [20, 25, 10, 10, 20]
    ]

    worksheet.write('A1', data)

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart2.add_series(:values => '=Sheet1!$B$1:$B$5')

    chart1.combine(chart2)

    # For testing
    chart1.instance_variable_set(:@cross_between, 'between')

    worksheet.insert_chart('E9', chart1)

    workbook.close
    compare_for_regression(
      [],
      { 'xl/charts/chart1.xml' => [
          '<c:dispBlanksAs'
        ] }
    )
  end
end
