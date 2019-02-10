# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartLine04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_line04
    @xlsx = 'chart_line04.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'line', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [47670016, 47671552])

    data = [
            [ 5,  2, 3, 4,  3 ],
            [ 10, 4, 6, 8,  6 ],
            [ 15, 6, 9, 12, 9 ]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5', :smooth => 1)
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5', :smooth => 1)

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
