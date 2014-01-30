# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_escapes07
    @xlsx = 'escapes07.xlsx'
    workbook   = WriteXLSX.new(@xlsx)
    worksheet  = workbook.add_worksheet

    worksheet.write_url(
                        'A1',
                        "http://example.com/!\"$%&'( )*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
                        )

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx
                                )
  end
end
