# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionProperties02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties02
    @xlsx = 'properties02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_properties(
      hyperlink_base: 'C:\\'
    )

    workbook.close
    compare_for_regression
  end
end
