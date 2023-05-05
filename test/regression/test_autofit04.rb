# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit04
    @xlsx = 'autofit04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write(0, 0, 'Hello')
    worksheet.write(0, 1, 'World')
    worksheet.write(0, 2, 123)
    worksheet.write(0, 3, 1234567)

    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
