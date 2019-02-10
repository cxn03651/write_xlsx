# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTabColor01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_tab_color01
    @xlsx = 'tab_color01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.tab_color = 'red'

    workbook.close
    compare_for_regression
  end
end
