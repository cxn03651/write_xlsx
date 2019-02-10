# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage13 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image13
    @xlsx = 'image13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(1, 75)
    worksheet.set_column('C:C', 32)

    worksheet.insert_image('C2',
                           'test/regression/images/logo.png', 8, 5)

    workbook.close
    compare_for_regression
  end
end
