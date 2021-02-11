# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink36 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink36
    @xlsx = 'hyperlink36.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'pie', :embedded => 1)

    worksheet.write('A1', 1)
    worksheet.write('A2', 2)

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      :url => 'https://github.com/jmcnamara'
    )

    chart.add_series(:values => '=Sheet1!$A$1:$A$2')
    worksheet.insert_chart('E12', chart)

    workbook.close

    compare_for_regression
  end
end
