# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_10
    @xlsx = 'utf8_10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'bar', :embedded => 1)

    chart.instance_variable_set('@axis_ids', [86604416, 89227648])

    data = [
            [ 1, 2, 3, 4,  5 ],
            [ 2, 4, 6, 8,  10 ],
            [ 3, 6, 9, 12, 15 ]
           ]

    worksheet.write( 'A1', data )

    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    chart.add_series(:values => '=Sheet1!$B$1:$B$5')
    chart.add_series(:values => '=Sheet1!$C$1:$C$5')

    chart.set_x_axis(:name => 'café')
    chart.set_y_axis(:name => 'sauté')
    chart.set_title(:name => 'résumé')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
