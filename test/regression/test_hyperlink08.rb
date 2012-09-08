# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_hyperlink08
    @xlsx = 'hyperlink08.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1',  'external://VBOXSVR/share/foo.xlsx', 'J:/foo.xlsx')
    worksheet.write_url('A3',  'external:foo.xlsx')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
