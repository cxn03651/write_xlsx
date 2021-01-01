# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_escapes07
    @xlsx = 'escapes07.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url(
                        'A1',
                        "http://example.com/!\"$%&'( )*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
                        )

    workbook.close
    compare_for_regression
  end
end
