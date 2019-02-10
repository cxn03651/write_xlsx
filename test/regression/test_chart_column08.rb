# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartColumn08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_column08
    @xlsx = 'chart_column08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [68809856, 68811392])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :categories      => '=(Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5)',
                     :values          => '=(Sheet1!$B$1:$B$2,Sheet1!$B$4:$B$5)',
                     :categories_data => [1, 2, 4, 5],
                     :values_data     => [2, 4, 8, 10]
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
