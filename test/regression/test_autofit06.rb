# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit06 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit06
    @xlsx = 'autofit06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_boolean(0, 0, true)
    worksheet.write_boolean(0, 1, false)

    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
