# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionObjectPosition08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position08
    @xlsx = 'object_position08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'line', :embedded => 1)

    bold   = workbook.add_format(:bold => 1)
    italic = workbook.add_format(:italic => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [60888960, 79670656])

    data = [
      [1, 2, 3,  4,  5],
      [2, 4, 6,  8, 10],
      [3, 6, 9, 12, 15]
    ]

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('B1', 'Bar', italic)
    worksheet.write('A2', data)

    worksheet.set_row(12, nil, nil, 1)
    worksheet.set_column('F:F', nil, nil, 1)

    chart.add_series(:values => '=Sheet1!$A$2:$A$6')
    chart.add_series(:values => '=Sheet1!$B$2:$B$6')
    chart.add_series(:values => '=Sheet1!$C$2:$C$6')

    worksheet.insert_chart('E9', chart, 0, 0, 1, 1, 2)

    workbook.close
    compare_for_regression
  end
end
