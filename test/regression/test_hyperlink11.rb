# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink11 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink11
    @xlsx = 'hyperlink11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(color: 'blue', underline: 1)

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1', 'http://www.perl.org/', format)

    workbook.close
    compare_for_regression
  end
end
