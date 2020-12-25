# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartPattern07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_pattern07
    @xlsx = 'chart_pattern07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [110902272, 110756608])

    data = [
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2],
      [2, 2, 2]
    ]

    worksheet.write('A1', data)

    chart.add_series(
      :values  => '=Sheet1!$A$1:$A$3',
      :pattern => {
        :pattern  => 'percent_40',
        :fg_color => '#C00000',
        :bg_color => '#FFFFFF'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$B$1:$B$3',
      :pattern => {
        :pattern  => 'percent_90',
        :fg_color => '#FF0000'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$C$1:$C$3',
      :pattern => {
        :pattern  => 'wide_upward_diagonal',
        :fg_color => '#FFC000'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$D$1:$D$3',
      :pattern => {
        :pattern  => 'dark_horizontal',
        :fg_color => '#FFFF00'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$E$1:$E$3',
      :pattern => {
        :pattern  => 'large_confetti',
        :fg_color => '#92D050'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$F$1:$F$3',
      :pattern => {
        :pattern  => 'plaid',
        :fg_color => '#00B050'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$G$1:$G$3',
      :pattern => {
        :pattern  => 'sphere',
        :fg_color => '#00B0F0'
      }
    )

    chart.add_series(
      :values  => '=Sheet1!$H$1:$H$3',
      :pattern => {
        :pattern  => 'solid_diamond',
        :fg_color => '#0070C0'
      }
    )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
