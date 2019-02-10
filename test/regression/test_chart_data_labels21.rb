# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartDataLabels21 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_data_labels21
    @xlsx = 'chart_data_labels21.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'pie', :embedded => 1)

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)

    chart.add_series(
                     :values      => '=Sheet1!$A$1:$A$5',
                     :data_labels => {
                       :value        => 1,
                       :category     => 1,
                       :series_name  => 1,
                       :percentage   => 1,
                       :separator    => ';',
                       :leader_lines => 1,
                       :position     => 'inside_end',
                       :legend_key   => 1,
                       :num_format   => '#,##0.00',
                       :font         => { :name => 'Consolas', :baseline => -1, :pitch_family => 49, :charset => 0 }
                     }
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
