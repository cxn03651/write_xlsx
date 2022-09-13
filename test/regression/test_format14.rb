# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionFormat14 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format14
    @xlsx = 'format14.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    center    = workbook.add_format

    center.set_center_across

    worksheet.write('A1', 'foo', center)

    workbook.close
    compare_for_regression
  end
end
