# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image02
    @xlsx = 'image02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('D7', 'test/regression/images/yellow.png', 1, 2)

    workbook.close
    compare_for_regression
  end
end
