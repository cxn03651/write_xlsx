# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink14 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink14
    @xlsx = 'hyperlink14.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.merge_range('C4:E5', 'http://www.perl.org/', format, 'Perl Home')

    workbook.close
    compare_for_regression
  end
end
