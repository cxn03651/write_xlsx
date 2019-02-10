# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRichString09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_rich_string09
    @xlsx = 'rich_string09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    bold   = workbook.add_format(:bold   => 1)
    italic = workbook.add_format(:italic => 1)

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('A2', 'Bar', italic)
    worksheet.write_rich_string('A3', 'a', bold, 'bc', 'defg')

    # The following contains 2 consectutive formats and should be ignored.
    worksheet.write_rich_string('A3', 'a', bold, bold, 'bc', 'defg')

    workbook.close
    compare_for_regression
  end
end
