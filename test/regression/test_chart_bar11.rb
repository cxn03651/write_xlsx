# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar11
    @xlsx = 'chart_bar11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart1    = workbook.add_chart(:type => 'bar', :embedded => 1)
    chart2    = workbook.add_chart(:type => 'bar', :embedded => 1)
    chart3    = workbook.add_chart(:type => 'bar', :embedded => 1)

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)
    worksheet.write('A7', 'http://www.perl.com/')
    worksheet.write('A8', 'http://www.perl.org/')
    worksheet.write('A9', 'http://www.perl.net/')

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart1.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart1.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart2.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart2.add_series(:values => '=Sheet1!$B$1:$B$5')

    chart3.add_series(:values => '=Sheet1!$A$1:$A$5')

    worksheet.insert_chart('E9',  chart1)
    worksheet.insert_chart('D25', chart2)
    worksheet.insert_chart('L32', chart3)

    workbook.close
    compare_for_regression(
      nil,
      {
        # Ignore the page margins.
        'xl/charts/chart1.xml' => [
          '<c:axId',
          '<c:crossAx',
          '<c:pageMargins'
        ],

        'xl/charts/chart2.xml' => [
          '<c:axId',
          '<c:crossAx',
          '<c:pageMargins'
        ],
        'xl/charts/chart3.xml' => [
          '<c:axId',
          '<c:crossAx',
          '<c:pageMargins'
        ]
      }
    )
  end
end
