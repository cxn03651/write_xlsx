# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar03
    @xlsx = 'chart_bar03.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart1      = workbook.add_chart(:type => 'bar', :embedded => 1)
    chart2      = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids, [64265216, 64447616])
    chart2.instance_variable_set(:@axis_ids, [86048128, 86058112])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet.write('A1', data)

    chart1.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$B$1:$B$5'
                     )
    chart1.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$C$1:$C$5'
                     )

    worksheet.insert_chart('E9', chart1)

    chart2.add_series(
                     :categories => '=Sheet1!$A$1:$A$4',
                     :values     => '=Sheet1!$B$1:$B$4'
                     )
    chart2.add_series(
                     :categories => '=Sheet1!$A$1:$A$4',
                     :values     => '=Sheet1!$C$1:$C$4'
                     )

    worksheet.insert_chart('F25', chart2)

    workbook.close
    compare_for_regression(
      nil,
      {
        # Ignore the page margins.
        'xl/charts/chart1.xml' => [ '<c:pageMargins' ],
        'xl/charts/chart2.xml' => [ '<c:pageMargins' ],
        # Ignore the workbookView.
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
