# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink29 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink29
    @xlsx = 'hyperlink29.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format1   = workbook.add_format(:hyperlink => 1)
    format2   = workbook.add_format(:color => 'red', :underline => 1)

    worksheet.write_url('A1', 'http://www.perl.org/', format1)
    worksheet.write_url('A2', 'http://www.perl.com/', format2)

    workbook.close

    compare_for_regression
  end
end
