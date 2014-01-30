# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_escapes08
    @xlsx = 'escapes08.xlsx'
    workbook   = WriteXLSX.new(@xlsx)
    worksheet  = workbook.add_worksheet

    # Test an already escaped string.
    worksheet.write_url(
                        'A1',
                        'http://example.com/%5b0%5d', 'http://example.com/[0]'
                        )

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx
                                )
  end
end
