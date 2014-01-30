# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_escapes02
    @xlsx = 'escapes02.xlsx'
    workbook    = WriteXLSX.new(@xlsx)
    worksheet   = workbook.add_worksheet

    worksheet.write('A1', %q{"<>'&})
    worksheet.write_comment('B2', %q{<>&"'})

    # Set the author to match the target XLSX file.
    worksheet.comments_author = %q{I am '"<>&}

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
