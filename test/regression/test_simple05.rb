# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSimple05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_simple05
    @xlsx = 'simple05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(5, 18)
    worksheet.set_row(6, 18)

    format1 = workbook.add_format(:bold      => 1)
    format2 = workbook.add_format(:italic    => 1)
    format3 = workbook.add_format(:bold      => 1, :italic => 1)
    format4 = workbook.add_format(:underline => 1)
    format5 = workbook.add_format(:font_strikeout => 1)
    format6 = workbook.add_format(:font_script => 1)
    format7 = workbook.add_format(:font_script => 2)

    worksheet.write_string(0, 0, 'Foo', format1)
    worksheet.write_string(1, 0, 'Foo', format2)
    worksheet.write_string(2, 0, 'Foo', format3)
    worksheet.write_string(3, 0, 'Foo', format4)
    worksheet.write_string(4, 0, 'Foo', format5)
    worksheet.write_string(5, 0, 'Foo', format6)
    worksheet.write_string(6, 0, 'Foo', format7)

    workbook.close
    compare_for_regression
  end
end
