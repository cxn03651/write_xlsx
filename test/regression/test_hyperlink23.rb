# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink23 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_hyperlink23
    @xlsx = 'hyperlink23.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1', 'https://en.wikipedia.org/wiki/Microsoft_Excel#Data_storage_and_communication', 'Display text')

    workbook.close

    compare_for_regression
  end
end
