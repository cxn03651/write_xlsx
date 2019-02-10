# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button02
    @xlsx = 'button02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('B4',
                            {
                              :x_offset => 4,
                              :y_offset => 3,
                              :caption  => 'my text'
                            }
                            )

    workbook.close
    compare_for_regression
  end
end
