# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink10
    @xlsx = 'hyperlink10.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:color => 'red', :underline => 1)

    worksheet.write_url('A1', 'http://www.perl.org/', format)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
