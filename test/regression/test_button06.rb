# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button06
    @xlsx = 'button05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('C2', {
                              :macro  => 'my_macro',
                              :width  => 128,
                              :height => 30
                            }
                            )

    workbook.close
    compare_for_regression
  end
end
