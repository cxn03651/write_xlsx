# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionEscapes08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_escapes08
    @xlsx = 'escapes08.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    # Test an already escaped string.
    worksheet.write_url(
      'A1',
      'http://example.com/%5b0%5d', 'http://example.com/[0]'
    )

    workbook.close
    compare_for_regression
  end
end
