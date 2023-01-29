# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionImage54 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image54
    @xlsx = 'image54.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      decorative: 1
    )

    workbook.close
    compare_for_regression
  end
end
