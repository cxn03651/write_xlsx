# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultFormat01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_default_format01
    @xlsx = 'default_format01.xlsx'
    workbook    = WriteXLSX.new(@xlsx, :default_format_properties => { :size => 10 })
    worksheet   = workbook.add_worksheet

    worksheet.set_default_row(12.75)

    # Override for testing
    worksheet.instance_variable_set(:@original_row_height, 12.75)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
