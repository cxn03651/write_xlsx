# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit08
    @xlsx = 'autofit08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string(0, 0, 'a')
    worksheet.write_string(1, 0, 'aaa')
    worksheet.write_string(2, 0, 'a')
    worksheet.write_string(3, 0, 'aaaa')
    worksheet.write_string(4, 0, 'a')

    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
