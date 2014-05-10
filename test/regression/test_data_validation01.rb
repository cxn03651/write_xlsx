# -*- coding: utf-8 -*-
require 'helper'

class TestDataValidation01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_data_validation01
    @xlsx = 'data_validation01.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.data_validation('C2', validate: 'list', value: ['Foo', 'Bar', 'Baz'])
    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
