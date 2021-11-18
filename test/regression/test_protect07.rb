# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProtect07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_protect07
    @xlsx = 'protect07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.read_only_recommended

    workbook.close
    compare_for_regression
  end
end
