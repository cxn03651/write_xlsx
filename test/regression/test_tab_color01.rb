# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTabColor01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_tab_color01
    @xlsx = 'tab_color01.xlsx'
    workbook    = WriteXLSX.new(@xlsx)
    worksheet   = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.tab_color = 'red'

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                nil,
                                nil
                                )
  end
end
