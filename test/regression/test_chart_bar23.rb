# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar23 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar23
    @xlsx = 'chart_bar23.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet
    chart      = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [43706240, 43727104])

    headers = ['Series 1', 'Series 2', 'Series 3']
    data = [
            [ 'Category 1', 'Category 2', 'Category 3', 'Category 4' ],
            [ 4.3,          2.5,          3.5,          4.5 ],
            [ 2.4,          4.5,          1.8,          2.8 ],
            [ 2,            2,            3,            5 ]
           ]

    worksheet.set_column('A:D', 11)

    worksheet.write('B1', headers)
    worksheet.write('A2', data)

    chart.add_series(
                     :categories      => '=Sheet1!$A$2:$A$5',
                     :values          => '=Sheet1!$B$2:$B$5'
                     )

    chart.add_series(
                     :categories      => '=Sheet1!$A$2:$A$5',
                     :values          => '=Sheet1!$C$2:$C$5'
                     )

    chart.add_series(
                     :categories      => '=Sheet1!$A$2:$A$5',
                     :values          => '=Sheet1!$D$2:$D$5'
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
      nil,
      {
        'xl/charts/chart1.xml' => ['<c:pageMargins']
      }
    )
  end
end
