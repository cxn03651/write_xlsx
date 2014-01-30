# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_escapes03
    @xlsx = 'escapes03.xlsx'
    workbook    = WriteXLSX.new(@xlsx)
    worksheet   = workbook.add_worksheet

    bold   = workbook.add_format(:bold   => 1)
    italic = workbook.add_format(:italic => 1)

    worksheet.write('A1', "Foo", bold)
    worksheet.write('A2', "Bar", italic)
    worksheet.write_rich_string('A3', 'a', bold, %q{b"<>'c}, 'defg')

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                {
                                  'xl/workbook.xml' => ['<workbookView']
                                }
                                )
  end
end
