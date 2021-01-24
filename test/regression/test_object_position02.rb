# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionObjectPosition02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position02
    @xlsx = 'object_position02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9',
      'test/regression/images/red.png',
      0, 0, 1, 1, 2
    )

    workbook.close
    compare_for_regression
  end
end
