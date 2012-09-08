# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink05
    @xlsx = 'hyperlink05.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1',  'http://www.perl.org/')
    worksheet.write_url('A3',  'http://www.perl.org/', 'Perl home')
    worksheet.write_url('A5',  'http://www.perl.org/', 'Perl home', nil, 'Tool Tip')
    worksheet.write_url('A7',  'http://www.cpan.org/', 'CPAN', nil, 'Download')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
