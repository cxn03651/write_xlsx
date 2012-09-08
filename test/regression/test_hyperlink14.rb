# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink14 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink14
    @xlsx = 'hyperlink14.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    worksheet.merge_range('C4:E5', 'http://www.perl.org/', format, 'Perl Home')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
