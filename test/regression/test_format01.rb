# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format01
    @xlsx = 'format01.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet('Data Sheet')
    worksheet3 = workbook.add_worksheet

    unused1 = workbook.add_format(:bold => 1)
    bold    = workbook.add_format(:bold => 1)
    unused2 = workbook.add_format(:bold => 1)
    unused3 = workbook.add_format(:italic => 1)

    worksheet1.write('A1', 'Foo')
    worksheet1.write('A2', 123)

    worksheet3.write('B2', 'Foo')
    worksheet3.write('B3', 'Bar', bold)
    worksheet3.write('C4', 234)

    workbook.close
    compare_for_regression
  end
end
