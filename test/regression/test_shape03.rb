# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShape03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shape03
    @xlsx = 'shape03.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    chart      = workbook.add_chart(:type => 'line', :embedded => 1)
    rect       = workbook.add_shape

    worksheet2.insert_shape('C2', rect)

    chart.instance_variable_set('@axis_ids', [110994176, 110996096])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet1.write('A1', data)

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    worksheet1.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
                                [],
                                {'xl/workbook.xml' => ['<workbookView']}
                                )
  end
end
