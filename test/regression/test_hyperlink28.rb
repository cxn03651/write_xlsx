# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHyperlink28 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink28
    @xlsx = 'hyperlink28.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:hyperlink => 1)

    worksheet.write_url('A1', 'http://www.perl.org/', format)

    workbook.close

    compare_for_regression
  end

  def test_hyperlink28_2
    @xlsx = 'hyperlink28.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1', 'http://www.perl.org/')

    workbook.close

    compare_for_regression
  end

  def test_hyperlink28_3
    @xlsx = 'hyperlink28.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.get_default_url_format

    worksheet.write_url('A1', 'http://www.perl.org/', format)

    workbook.close

    compare_for_regression
  end
end
