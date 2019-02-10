# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image04
    @xlsx = 'image04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/red.bmp')

    workbook.close
    compare_for_regression
  end
end
