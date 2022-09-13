# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink21 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink21
    @xlsx = 'hyperlink21.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1', 'external:C:\Temp\Test 1')

    workbook.close

    compare_for_regression
  end
end
