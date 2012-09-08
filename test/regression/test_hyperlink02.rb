# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink02
    @xlsx = 'hyperlink02.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'http://www.perl.org/')
    worksheet.write('D4', 'http://www.perl.org/')
    worksheet.write('A8', 'http://www.perl.org/')
    worksheet.write('B6', 'http://www.cpan.org/')
    worksheet.write('F12', 'http://www.cpan.org/')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
