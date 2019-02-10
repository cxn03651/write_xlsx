# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartBar18 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_bar18
    @xlsx = 'chart_bar18.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet
    chart      = workbook.add_chart(:type => 'bar')

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [40294272, 40295808])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart.activate
    chart.set_header('Page &P')
    chart.set_footer('&A')

    workbook.close
    compare_for_regression(
      [
        'xl/printerSettings/printerSettings1.bin',
        'xl/chartsheets/_rels/sheet1.xml.rels',
        '[Content_Types].xml'
      ],
      {
        'xl/chartsheets/sheet1.xml' => [
          '<pageMargins',
          '<pageSetup',
          '<drawing',    # Id is wrong due to missing printerbin.
        ],
        'xl/charts/chart1.xml' => ['<c:pageMargins']
      }
    )
  end
end
