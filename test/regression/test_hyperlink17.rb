# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink17 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink17
    @xlsx = 'hyperlink17.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    # Test URL with whitespace.
    worksheet.write_url('A1', 'http://google.com/some link')

    workbook.close
    compare_for_regression(
                                {},
                                { 'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
