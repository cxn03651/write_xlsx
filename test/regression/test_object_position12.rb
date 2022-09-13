# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionObjectPosition12 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_object_position12
    @xlsx = 'object_position12.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column(1, 1, 5, nil)

    worksheet.insert_image('E9', 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
