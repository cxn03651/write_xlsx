# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTypes02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_types02
    @xlsx = 'types02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_boolean('A1', 1)
    worksheet.write_boolean('A2', false)

    workbook.close

    compare_for_regression
  end
end
