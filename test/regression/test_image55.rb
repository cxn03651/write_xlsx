# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage55 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image55
    @xlsx = 'image55.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      :url        => 'https://github.com/jmcnamara',
      :decorative => 1
    )

    workbook.close
    compare_for_regression
  end
end
