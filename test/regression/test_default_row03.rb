# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_default_row03
    @xlsx = 'default_row03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(24, 1)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')

    (1..8).each { |row| worksheet.set_row(row, 24) }

    workbook.close
    compare_for_regression
  end
end
