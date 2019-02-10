# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage16 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image16
    @xlsx = 'image16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('C2',
                           'test/regression/images/issue32.png')

    workbook.close
    compare_for_regression
  end
end
