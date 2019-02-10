# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button05
    @xlsx = 'button05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('C2', {
                              :macro   => 'my_macro',
                              :x_scale => 2,
                              :y_scale => 1.5
                            }
                            )

    workbook.close
    compare_for_regression
  end
end
