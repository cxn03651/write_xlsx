# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink35 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink35
    @xlsx = 'hyperlink35.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'A1', 'test/regression/images/blue.png',
      url: 'https://github.com/foo'
    )
    worksheet.insert_image(
      'B3', 'test/regression/images/red.jpg',
      url: 'https://github.com/bar'
    )
    worksheet.insert_image(
      'D5', 'test/regression/images/yellow.jpg',
      url: 'https://github.com/baz'
    )
    worksheet.insert_image(
      'F9', 'test/regression/images/grey.png',
      url: 'https://github.com/boo'
    )

    workbook.close

    compare_for_regression
  end
end
