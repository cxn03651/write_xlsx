# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink12
    @xlsx = 'hyperlink12.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:color => 'blue', :underline => 1)

    worksheet.write_url('A1', 'mailto:jmcnamara@cpan.org', format)
    worksheet.write_url('A3', 'ftp://perl.org/', format)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
