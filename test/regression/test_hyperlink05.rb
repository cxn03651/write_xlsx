# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink05
    @xlsx = 'hyperlink05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Turn off default URL format for testing.
    worksheet.instance_variable_set(:@default_url_format, nil)

    worksheet.write_url('A1',  'http://www.perl.org/')
    worksheet.write_url('A3',  'http://www.perl.org/', 'Perl home')
    worksheet.write_url('A5',  'http://www.perl.org/', 'Perl home', nil, 'Tool Tip')
    worksheet.write_url('A7',  'http://www.cpan.org/', 'CPAN', nil, 'Download')

    workbook.close
    compare_for_regression
  end
end
