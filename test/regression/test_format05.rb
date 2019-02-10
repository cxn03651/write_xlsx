# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format05
    @xlsx = 'format05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    wrap      = workbook.add_format(:text_wrap => 1)

    worksheet.set_row(0, 45)

    worksheet.write('A1', "Foo\nBar", wrap)

    workbook.close
    compare_for_regression
  end
end
