# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit10 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit10
    @xlsx = 'autofit10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    bold = workbook.add_format(bold: 1)

    worksheet.write_rich_string(0, 0, "F", bold, "o", "o", bold, "b", "a", bold, 'r')
    worksheet.write(1, 0, "Bar", bold)

    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
