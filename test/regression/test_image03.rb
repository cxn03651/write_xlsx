# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image03
    @xlsx = 'image03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/red.jpg')

    workbook.close
    compare_for_regression
  end
end
