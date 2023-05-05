# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit09 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit09
    @xlsx = 'autofit09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    text_wrap = workbook.add_format(text_wrap: 1)

    worksheet.write_string(0, 0, "Hello\nFoo", text_wrap)
    worksheet.write_string(2, 2, "Foo\nBamboo\nBar", text_wrap)

    worksheet.set_row(0, 33)
    worksheet.set_row(2, 48)

    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
