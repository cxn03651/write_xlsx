# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionObjectPosition16 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position16
    @xlsx = 'object_position16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column(1, 1, nil, nil, 1)

    # Same as testcase test_object_position13 except with an offset.
    worksheet.insert_image(
      'A9', 'test/regression/images/red.png',
      :x_offset => 192
    )

    workbook.close
    compare_for_regression
  end
end
