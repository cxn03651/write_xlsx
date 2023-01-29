# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink48 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink48
    @xlsx = 'hyperlink48.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      url: 'https://github.com/jmcnamara'
    )
    worksheet.insert_image(
      'E13', 'test/regression/images/red.png',
      url: 'https://github.com/jmcnamara'
    )

    workbook.close

    compare_for_regression
  end
end
