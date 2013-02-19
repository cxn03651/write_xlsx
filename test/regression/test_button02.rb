# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_button02
    @xlsx = 'button02.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('B4',
                            {
                              :x_offset => 4,
                              :y_offset => 3,
                              :caption  => 'my text'
                            }
                            )

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
