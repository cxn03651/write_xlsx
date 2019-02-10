# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar02
    @xlsx = 'chart_bar02.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet
    worksheet2  = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [93218304, 93219840])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet1.write('A1', 'Foo')
    worksheet2.write('A1', data)

    chart.add_series(
                     :categories => '=Sheet2!$A$1:$A$5',
                     :values     => '=Sheet2!$B$1:$B$5'
                     )
    chart.add_series(
                     :categories => '=Sheet2!$A$1:$A$5',
                     :values     => '=Sheet2!$C$1:$C$5'
                     )

    worksheet2.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
      nil,
      {
        # Ignore the page margins.
        'xl/charts/chart1.xml' => [ '<c:pageMargins' ],
        # Ignore the workbookView.
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
