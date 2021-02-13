# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionObjectPosition13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position13
    @xlsx = 'object_position13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column(1, 1, nil, nil, 1)

    worksheet.insert_image('E9', 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
