# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.coom
#

require 'helper'

class TestRegressionAutofit05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_autofit05
    @xlsx = 'autofit05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    date_format = workbook.add_format(num_format: 14)

    worksheet.write_date_time(0, 0, '2023-01-01T', date_format)
    worksheet.write_date_time(0, 1, '2023-12-12T', date_format)

    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
