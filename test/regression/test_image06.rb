# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image06
    @xlsx = 'image06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'bar', :embedded => 1)

    # For testig, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [46335488, 46364544])

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', data)
    chart.add_series(:values => '=Sheet1!$A$1:$A$5')
    worksheet.insert_chart('E9', chart)
    worksheet.insert_image('F2', 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
