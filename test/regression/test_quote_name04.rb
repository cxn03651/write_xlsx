# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionQuoteName04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_quote_name04
    @xlsx = 'quote_name04.xlsx'
    workbook  = WriteXLSX.new(@io)

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    # Test quoted/non-quoted sheet names.
    worksheet = workbook.add_worksheet('Sheet 1')
    chart = workbook.add_chart(:type => 'pie', :embedded => 1)

    worksheet.write( 'A1', data )
    chart.add_series(
      :values => ['Sheet 1', 0, 4, 0, 0],
      :name   => ['Sheet 1', 0, 0]
    )

    chart.set_title(:name => 'Foo')

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression
  end
end
