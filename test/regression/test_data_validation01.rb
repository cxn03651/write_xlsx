# -*- coding: utf-8 -*-
require 'helper'

class TestDataValidation01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_data_validation01
    @xlsx = 'data_validation01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.data_validation('C2', validate: 'list', value: ['Foo', 'Bar', 'Baz'])
    workbook.close
    compare_for_regression
  end
end
