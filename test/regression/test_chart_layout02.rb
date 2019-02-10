# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartLayout02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_layout02
    @xlsx = 'chart_layout02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [68311296, 69198208])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart.set_legend(
                       :layout => {
                         :x      => 0.80197353455818021,
                         :y      => 0.37442403032954213,
                         :width  => 0.12858202099737534,
                         :height => 0.25115157480314959
                       }
                       )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
