# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartsheet08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chartsheet08
    @xlsx = 'chartsheet08.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'bar')

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids,  [46320256, 46335872])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    # Chartsheet test.
    chart.margin_left   = '0.70866141732283472'
    chart.margin_right  = '0.70866141732283472'
    chart.margin_top    = '0.74803149606299213'
    chart.margin_bottom = '0.74803149606299213'
    chart.set_header('Page &P', '0.51181102362204722')
    chart.set_footer('&A'     , '0.51181102362204722')

    workbook.close
    compare_for_regression(
      %w[
        xl/printerSettings/printerSettings1.bin
        xl/chartsheets/_rels/sheet1.xml.rels
      ],
      {
        '[Content_Types].xml'       => ['<Default Extension="bin"'],
        'xl/workbook.xml'           => ['<workbookView'],
        'xl/chartsheets/sheet1.xml' => [
          '<pageSetup',
          '<drawing',    # Id is wrong due to missing printerbin.
        ]
      }
    )
  end
end
