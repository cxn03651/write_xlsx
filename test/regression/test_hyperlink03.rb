# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink03
    @xlsx = 'hyperlink03.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet1.instance_variable_set(:@default_url_format, nil)
    worksheet2.instance_variable_set(:@default_url_format, nil)

    worksheet1.write_url('A1', 'http://www.perl.org/')
    worksheet1.write_url('D4', 'http://www.perl.org/')
    worksheet1.write_url('A8', 'http://www.perl.org/')
    worksheet1.write_url('B6', 'http://www.cpan.org/')
    worksheet1.write_url('F12', 'http://www.cpan.org/')

    worksheet2.write_url('C2', 'http://www.google.com/')
    worksheet2.write_url('C5', 'http://www.cpan.org/')
    worksheet2.write_url('C7', 'http://www.perl.org/')

    workbook.close
    compare_for_regression
  end
end
