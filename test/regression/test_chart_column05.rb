# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartColumn05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_column05
    @xlsx = 'chart_column05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet('Foo')
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids,  [47292800, 47295104])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Foo!$A$1:$A$5')
    chart.add_series(:values => '=Foo!$B$1:$B$5')
    chart.add_series(:values => '=Foo!$C$1:$C$5')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
