# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image09
    @xlsx = 'image09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9',
                           'test/regression/images/red_64x20.png')

    workbook.close
    compare_for_regression
  end
end
