# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_default_row01
    @xlsx = 'default_row01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(24)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')

    workbook.close
    compare_for_regression
  end
end
