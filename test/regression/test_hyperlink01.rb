# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink01
    @xlsx = 'hyperlink01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write('A1', 'http://www.perl.org/')

    workbook.close
    compare_for_regression
  end
end
