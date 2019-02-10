# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartDate02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_date02
    @xlsx = 'chart_date02.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'line', :embedded => 1)
    date_format = workbook.add_format(:num_format => 14)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [55112064, 55115136])

    worksheet.set_column('A:A', 12)

    dates = [
             '2013-01-01T', '2013-01-02T', '2013-01-03T', '2013-01-04T',
             '2013-01-05T', '2013-01-06T', '2013-01-07T', '2013-01-08T',
             '2013-01-09T', '2013-01-10T'
            ]

    data = [10, 30, 20, 40, 20, 60, 50, 40, 30, 30]

    dates.each_with_index do |date_time, row|
      worksheet.write_date_time(row, 0, date_time, date_format)
      worksheet.write(row, 1, data[row])
    end

    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$10',
                     :values     => '=Sheet1!$B$1:$B$10'
                     )

    chart.set_x_axis(
                     :date_axis         => 1,
                     :min               => worksheet.convert_date_time('2013-01-02T'),
                     :max               => worksheet.convert_date_time('2013-01-09T'),
                     :num_format        => 'dd/mm/yyyy',
                     :num_format_linked => 1,
                     )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
                                nil,
                                {'xl/charts/chart1.xml' => ['<c:formatCode']}
                                )
  end
end
