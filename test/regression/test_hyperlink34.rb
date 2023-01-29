# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink34 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink34
    @xlsx = 'hyperlink34.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('A1', 'test/regression/images/blue.png')
    worksheet.insert_image(
      'B3', 'test/regression/images/red.jpg',
      url: 'https://github.com/jmcnamara'
    )
    worksheet.insert_image('D5', 'test/regression/images/yellow.jpg')
    worksheet.insert_image(
      'F9', 'test/regression/images/grey.png',
      url: 'https://github.com'
    )

    workbook.close

    compare_for_regression
  end
end
