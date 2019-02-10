# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar04
    @xlsx = 'chart_bar04.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet
    worksheet2  = workbook.add_worksheet
    chart1      = workbook.add_chart(:type => 'bar', :embedded => 1)
    chart2      = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids, [64446848, 64448384])
    chart2.instance_variable_set(:@axis_ids, [85389696, 85391232])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet1.write('A1', data)

    chart1.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$B$1:$B$5'
                     )
    chart1.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$C$1:$C$5'
                     )

    worksheet1.insert_chart('E9', chart1)

    worksheet2.write('A1', data)

    chart2.add_series(
                     :categories => '=Sheet2!$A$1:$A$5',
                     :values     => '=Sheet2!$B$1:$B$5'
                     )
    chart2.add_series(
                     :categories => '=Sheet2!$A$1:$A$5',
                     :values     => '=Sheet2!$C$1:$C$5'
                     )

    worksheet2.insert_chart('E9', chart2)

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
