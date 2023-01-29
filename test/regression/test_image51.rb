# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage51 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image51
    @xlsx = 'image51.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9',  'test/regression/images/red.png',
      url: 'https://duckduckgo.com/?q=1'
    )
    worksheet.insert_image(
      'E13', 'test/regression/images/red2.png',
      url: 'https://duckduckgo.com/?q=2'
    )

    workbook.close
    compare_for_regression
  end
end
