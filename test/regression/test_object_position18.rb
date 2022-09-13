# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionObjectPosition18 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position18
    @xlsx = 'object_position18.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column(1, 1, 5, nil)

    # Same as testcase test_object_position12 except with an offset.
    worksheet.insert_image(
      'B9', 'test/regression/images/red.png',
      :x_offset => 168
    )

    workbook.close
    compare_for_regression
  end
end
