# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSharedStrings02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shared_strings02
    @xlsx = 'shared_strings02.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    # Test that the Excel/Unicode escape for control characters, _xHHHH_ is
    # also escaped when written as a literal string.

    strings = [
               "_",
               "_x",
               "_x0",
               "_x00",
               "_x000",
               "_x0000",
               "_x0000_",
               "_x005F_",
               "_x000G_",
               "_X0000_",
               "_x000a_",
               "_x000A_",
               "_x0000__x0000_",
               "__x0000__"
              ]
    worksheet.write_col(0, 0, strings)

    workbook.close
    compare_for_regression(
      nil,
      {'xl/workbook.xml' => ['<workbookView']}
    )
  end
end
