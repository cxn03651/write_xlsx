# -*- coding: utf-8 -*-

#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to Ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'helper'

class TestRegressionProtect08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_protect08
    @xlsx = 'protect08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    options = {
      objects:               1,
      scenarios:             1,
      format_cells:          1,
      format_columns:        1,
      format_rows:           1,
      insert_columns:        1,
      insert_rows:           1,
      insert_hyperlinks:     1,
      delete_columns:        1,
      delete_rows:           1,
      select_locked_cells:   0,
      sort:                  1,
      autofilter:            1,
      pivot_tables:          1,
      select_unlocked_cells: 0
    }

    worksheet.protect("", options)

    workbook.close
    compare_for_regression
  end
end
