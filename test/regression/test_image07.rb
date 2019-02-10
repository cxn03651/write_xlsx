# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image07
    @xlsx = 'image07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_image('E9', 'test/regression/images/red.png')
    worksheet2.insert_image('E9', 'test/regression/images/yellow.png')

    workbook.close
    compare_for_regression
  end
end
