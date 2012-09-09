# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_image07
    @xlsx = 'image07.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_image('E9', 'test/regression/images/red.png')
    worksheet2.insert_image('E9', 'test/regression/images/yellow.png')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
