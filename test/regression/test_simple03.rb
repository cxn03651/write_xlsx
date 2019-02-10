# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSimple03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_simple03
    @xlsx = 'simple03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet('Data Sheet')
    worksheet3 = workbook.add_worksheet

    bold = workbook.add_format(:bold => 1)

    worksheet1.write('A1', 'Foo')
    worksheet1.write('A2', 123)

    worksheet3.write('B2', 'Foo')
    worksheet3.write('B3', 'Bar', bold)
    worksheet3.write('C4', 234)

    # This should be overridden by the worksheet3 call below.
    worksheet2.activate

    worksheet2.select
    worksheet3.select
    worksheet3.activate

    workbook.close
    compare_for_regression
  end
end
