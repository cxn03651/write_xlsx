# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionObjectPosition17 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position17
    @xlsx = 'object_position17.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column(1, 1, 5, nil, 1)

    # Same as testcase test_object_position14 except with an offset.
    worksheet.insert_image(
      'A9', 'test/regression/images/red.png',
      x_offset: 192
    )

    workbook.close
    compare_for_regression
  end
end
