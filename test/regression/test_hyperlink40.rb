# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink40 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink40
    @xlsx = 'hyperlink40.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      url: 'https://github.com/jmcnamara#foo'
    )

    workbook.close

    compare_for_regression
  end
end
