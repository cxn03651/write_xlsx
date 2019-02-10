# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image10
    @xlsx = 'image10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('C2',
                           'test/regression/images/logo.png')

    workbook.close
    compare_for_regression
  end
end
