# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image11
    @xlsx = 'image11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('C2',
                           'test/regression/images/logo.png', 8, 5)

    workbook.close
    compare_for_regression
  end
end
