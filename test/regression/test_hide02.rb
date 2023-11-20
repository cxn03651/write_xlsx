# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHide02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hide02
    @xlsx = 'hide02.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet2.very_hidden

    workbook.close
    compare_for_regression
  end
end
