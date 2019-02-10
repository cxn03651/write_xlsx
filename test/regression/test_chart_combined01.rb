# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartCombined01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_combined01
    @xlsx = 'chart_combined01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart1    = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart1.instance_variable_set(:@axis_ids, [84882560, 84884096])

    data = [
      [  2,  7,  3,  6,   2 ],
      [ 20, 25, 10, 10,  20 ]
    ]

    worksheet.write('A1', data)

    chart1.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart1.add_series(:values => '=Sheet1!$B$1:$B$5')

    worksheet.insert_chart('E9', chart1)

    workbook.close
    compare_for_regression
  end
end
