# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink02
    @xlsx = 'hyperlink02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'http://www.perl.org/')
    worksheet.write('D4', 'http://www.perl.org/')
    worksheet.write('A8', 'http://www.perl.org/')
    worksheet.write('B6', 'http://www.cpan.org/')
    worksheet.write('F12', 'http://www.cpan.org/')

    workbook.close
    compare_for_regression
  end
end
