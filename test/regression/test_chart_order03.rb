# -*- coding: utf-8 -*-
require 'helper'

class TestChartOrder03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_order03
    @xlsx = 'chart_order03.xlsx'
    workbook    =  WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet
    worksheet2  = workbook.add_worksheet
    chart2      = workbook.add_chart(:type => 'bar')
    worksheet3  = workbook.add_worksheet

    chart4      = workbook.add_chart(:type => 'pie',    :embedded => 1)
    chart3      = workbook.add_chart(:type => 'line',   :embedded => 1)
    chart1      = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xls file.
    chart1.instance_variable_set(:@axis_ids, [67913600, 68169088])
    chart2.instance_variable_get(:@chart).
      instance_variable_set(:@axis_ids, [58117120, 67654400])
    chart3.instance_variable_set(:@axis_ids, [58109952, 68215936])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet1.write('A1', data)
    worksheet2.write('A1', data)
    worksheet3.write('A1', data)

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart2.add_series(:values => '=Sheet2!$A$1:$A$5')
    chart3.add_series(:values => '=Sheet3!$A$1:$A$5')
    chart4.add_series(:values => '=Sheet1!$B$1:$B$5')

    worksheet1.insert_chart('E9',  chart1)
    worksheet3.insert_chart('E9',  chart3)
    worksheet1.insert_chart('E24', chart4)

    workbook.close
    compare_for_regression(
      [],
      {
        'xl/charts/chart1.xml' => [ '<c:formatCode', '<c:pageMargins' ],
        'xl/charts/chart2.xml' => [ '<c:formatCode', '<c:pageMargins' ],
        'xl/charts/chart4.xml' => [ '<c:formatCode', '<c:pageMargins' ]
      }
    )
  end
end
