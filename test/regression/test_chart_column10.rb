# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartColumn10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_chart_column10
    @xlsx = 'chart_column10.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [45686144, 45722240 ])

    data = [
            [ 'A', 'B', 'C', 'D', 'E' ],
            [  1,   2,   3,   2,   1  ]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :categories      => 'Sheet1!$A$1:$A$5',
                     :values          => 'Sheet1!$B$1:$B$5'
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx),
                                @xlsx,
                                nil,
                                nil
                                )
  end
end
