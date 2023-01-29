# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink49 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink49
    @xlsx = 'hyperlink49.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1', 'https://github.com/jmcnamara')

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      url: 'https://github.com/jmcnamara'
    )

    workbook.close

    compare_for_regression
  end
end
