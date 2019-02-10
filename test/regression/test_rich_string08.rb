# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRichString08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_rich_string08
    @xlsx = 'rich_string08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    bold   = workbook.add_format(:bold   => 1)
    italic = workbook.add_format(:italic => 1)
    format = workbook.add_format(:align  => 'center')

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('A2', 'Bar', italic)
    worksheet.write_rich_string('A3', 'ab', bold, 'cd', 'efg', format)

    workbook.close
    compare_for_regression
  end
end
