# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultFormat01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_default_format01
    @xlsx = 'default_format01.xlsx'
    workbook    = WriteXLSX.new(@io, :default_format_properties => { :size => 10 })
    worksheet   = workbook.add_worksheet

    worksheet.set_default_row(12.75)

    # Override for testing
    worksheet.instance_variable_set(:@original_row_height, 12.75)

    workbook.close
    compare_for_regression
  end
end
