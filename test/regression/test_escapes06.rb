# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_escapes06
    @xlsx = 'escapes06.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet
    num_format = workbook.add_format(:num_format => '[Red]0.0%\\ "a"')

    worksheet.set_column('A:A', 14)

    worksheet.write('A1', 123, num_format)

    workbook.close
    compare_for_regression
  end
end
