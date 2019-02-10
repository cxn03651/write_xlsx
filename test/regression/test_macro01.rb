# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionMacro01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_macro01
    @xlsx = 'macro01.xlsm'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.add_vba_project(File.join(
                                       @regression_output,
                                       'vbaProject01.bin'
                                       )
                             )

    worksheet.write('A1', 123)

    workbook.close
    compare_for_regression
  end
end
