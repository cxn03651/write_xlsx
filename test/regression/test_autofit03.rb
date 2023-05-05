# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit03
    @xlsx = 'autofit03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string(0, 0, 'A')
    worksheet.write_string(0, 1, 'A')

    worksheet.autofit

    worksheet.set_column('B:B', 1.57)

    workbook.close
    compare_for_regression
  end
end
