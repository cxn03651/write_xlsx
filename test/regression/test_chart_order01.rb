# -*- coding: utf-8 -*-
require 'helper'

class TestChartOrder01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_order01
    @xlsx = 'chart_order01.xlsx'
    workbook    =  WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet
    worksheet2  = workbook.add_worksheet
    worksheet3  = workbook.add_worksheet

    chart1      = workbook.add_chart(:type => 'column', :embedded => 1)
    chart2      = workbook.add_chart(:type => 'bar',    :embedded => 1)
    chart3      = workbook.add_chart(:type => 'line',   :embedded => 1)
    chart4      = workbook.add_chart(:type => 'pie',    :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xls file.
    chart1.instance_variable_set(:@axis_ids, [54976896, 54978432])
    chart2.instance_variable_set(:@axis_ids, [54310784, 54312320])
    chart3.instance_variable_set(:@axis_ids, [69816704, 69818240])
    chart4.instance_variable_set(:@axis_ids, [69816704, 69818240])

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
    worksheet2.insert_chart('E9',  chart2)
    worksheet3.insert_chart('E9',  chart3)
    worksheet1.insert_chart('E24', chart4)

    workbook.close
    compare_for_regression
  end
end
