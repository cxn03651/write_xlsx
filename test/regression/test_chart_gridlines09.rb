# -*- coding: utf-8 -*-
require 'helper'

class TestChartGridlines09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_gridlines09
    @xlsx = 'chart_gridlines09.xlsx'
    workbook    =  WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xls file.
    chart.instance_variable_set(:@axis_ids, [48744320, 49566848])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ],
           ]

    worksheet.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart.set_x_axis(
                     :major_gridlines => {
                       :visible => 1,
                       :line    => {
                         :color     => 'red',
                         :width     => 0.5,
                         :dash_type => 'square_dot'
                       }
                     },
                     :minor_gridlines => {
                       :visible => 1,
                       :line    => { :color => 'yellow' }
                     }
                     )
    chart.set_y_axis(
                     :major_gridlines => {
                       :visible => 1,
                       :line    => {
                         :width     => 1.25,
                         :dash_type => 'dash'
                       }
                     },
                     :minor_gridlines => {
                       :visible => 1,
                       :line    => { :color => '#00B050' }
                     }
                     )

    worksheet.insert_chart('E9',  chart)

    workbook.close
    compare_for_regression
  end
end
