# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_excel2003_style08
    @xlsx = 'excel2003_style08.xlsx'
    workbook    = WriteXLSX.new(@io, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    courier = workbook.add_format(:font => 'Courier', :size => 8, :font_family => 3)

    worksheet.write('A1', 'Foo')
    worksheet.write('A2', 'Bar', courier)

    workbook.close
    compare_for_regression
  end
end
