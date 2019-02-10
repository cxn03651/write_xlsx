# -*- coding: utf-8 -*-
require 'helper'

class TestChartGridlines07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_gridlines07
    @xlsx = 'chart_gridlines07.xlsx'
    workbook    =  WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    chart       = workbook.add_chart(:type => 'stock', :embedded => 1)
    date_format = workbook.add_format(:num_format => 14)

    # For testing, copy the randomly generated axis ids in the target xls file.
    chart.instance_variable_set(:@axis_ids, [59313152, 59364096])

    data = [
            [ '2007-01-01T', '2007-01-02T', '2007-01-03T', '2007-01-04T', '2007-01-05T' ],
            [ 27.2,  25.03, 19.05, 20.34, 18.5 ],
            [ 23.49, 19.55, 15.12, 17.84, 16.34 ],
            [ 25.45, 23.05, 17.32, 20.45, 17.34 ]
           ]

    (0..4).each do |row|
      worksheet.write_date_time(row, 0, data[0][row], date_format)
      worksheet.write(row, 1, data[1][row])
      worksheet.write(row, 2, data[2][row])
      worksheet.write(row, 3, data[3][row])
    end

    worksheet.set_column('A:D', 11)

    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$B$1:$B$5'
                     )
    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$C$1:$C$5'
                     )
    chart.add_series(
                     :categories => '=Sheet1!$A$1:$A$5',
                     :values     => '=Sheet1!$D$1:$D$5'
                     )

    chart.set_x_axis(
                     :major_gridlines => { :visible => 1 },
                     :minor_gridlines => { :visible => 1 }
                     )
    chart.set_y_axis(
                     :major_gridlines => { :visible => 1 },
                     :minor_gridlines => { :visible => 1 }
                     )

    worksheet.insert_chart('E9',  chart)

    workbook.close
    compare_for_regression(
      [],
      {'xl/charts/chart1.xml' => ['<c:formatCode']}
    )
  end
end
