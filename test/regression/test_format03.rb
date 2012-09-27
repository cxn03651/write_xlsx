# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_format03
    @xlsx = 'format03.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(:bold     => 1, :fg_color => 'red')
    format2 = workbook.add_format

    # Test the copy method
    format2.copy(format1)
    format2.set_italic

    worksheet.write('A1', 'Foo', format1)
    worksheet.write('A2', 'Bar', format2)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
