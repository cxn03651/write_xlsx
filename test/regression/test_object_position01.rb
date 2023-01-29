# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionObjectPosition01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position01
    @xlsx = 'object_position01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image(
      'E9', 'test/regression/images/red.png',
      object_position: 1
    )

    workbook.close
    compare_for_regression
  end
end
