# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink10
    @xlsx = 'hyperlink10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:color => 'red', :underline => 1)

    worksheet.write_url('A1', 'http://www.perl.org/', format)

    workbook.close
    compare_for_regression
  end
end
