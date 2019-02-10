# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRichString10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_rich_string10
    @xlsx = 'rich_string10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    bold   = workbook.add_format(:bold   => 1)
    italic = workbook.add_format(:italic => 1)

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('A2', 'Bar', italic)

    # Verify that whitespace is preserved.
    worksheet.write_rich_string('A3', ' a', bold, 'bc', 'defg ')

    workbook.close
    compare_for_regression
  end
end
