# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSharedStrings01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_shared_strings01
    @xlsx = 'shared_strings01.xlsx'
    workbook    = WriteXLSX.new(@xlsx)
    worksheet   = workbook.add_worksheet

    # Test that control characters and any other single byte characters are
    # handled correctly by the SharedStrings module. We skip chr 34 = " in
    # this test since it isn't encode by Excel as &quot;.
    (0..255).each do |i|
      next if i == 34
      worksheet.write_string(i, 0, i.chr)
    end

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                nil,
                                {'xl/workbook.xml' => ['<workbookView']}
                                )
  end
end
