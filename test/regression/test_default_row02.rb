# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_default_row02
    @xlsx = 'default_row02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(nil, 1)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')

    (1..8).each { |row| worksheet.set_row(row, 15) }

    workbook.close
    compare_for_regression
  end
end
