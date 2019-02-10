# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink13 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink13
    @xlsx = 'hyperlink13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    worksheet.merge_range('C4:E5', 'http://www.perl.org/', format)

    workbook.close
    compare_for_regression
  end
end
