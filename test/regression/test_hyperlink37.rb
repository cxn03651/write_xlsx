# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink37 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink37
    @xlsx = 'hyperlink37.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    textbox = workbook.add_shape

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      :url => 'https://github.com/jmcnamara'
    )

    worksheet.insert_shape('E12', textbox)

    workbook.close

    compare_for_regression(
      ['xl/drawings/drawing1.xml'], {}
    )
  end
end
