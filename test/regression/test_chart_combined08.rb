# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartCombined08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_combined08
    @xlsx = 'chart_combined08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart1    = workbook.add_chart(:type => 'column',  :embedded => 1)
    chart2    = workbook.add_chart(:type => 'scatter', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids,  [81267328, 81297792])
    chart2.instance_variable_set(:@axis_ids,  [81267328, 81297792])
    chart2.instance_variable_set(:@axis2_ids, [89510656, 84556032])

    data = [
      [ 2,  3,  4,  5,  6],
      [20, 25, 10, 10, 20],
      [ 5, 10, 15, 10,  5]
    ]

    worksheet.write('A1', data)

    chart1.add_series(
      :categories => '=Sheet1!$A$1:$A$5',
      :values     => '=Sheet1!$B$1:$B$5'
    )
    chart2.add_series(
      :categories => '=Sheet1!$A$1:$A$5',
      :values     => '=Sheet1!$C$1:$C$5',
      :y2_axis    => 1
    )

    chart1.combine(chart2)

    worksheet.insert_chart('E9', chart1)

    workbook.close
    compare_for_regression(
      [],
      { 'xl/charts/chart1.xml' => [
          '<c:dispBlanksAs',
          '<c:crossBetween',
          '<c:tickLblPos',
          '<c:auto',
          '<c:valAx>',
          '<c:catAx>',
          '</c:valAx>',
          '</c:catAx>',
          '<c:crosses',
          '<c:lblOffset',
          '<c:lblAlgn'
        ] }
    )
  end
end
