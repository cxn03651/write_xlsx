# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format13
    @xlsx = 'format13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(0, 21)

    font_format = workbook.add_format

    font_format.set_font('B Nazanin')
    font_format.set_font_family(0)
    font_format.set_font_charset(178)

    worksheet.write('A1', 'Foo', font_format)

    workbook.close
    compare_for_regression
  end
end
