# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button07
    @xlsx = 'button07.xlsm'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_vba_name('ThisWorkbook')
    worksheet.set_vba_name('Sheet1')

    workbook.set_vba_name
    worksheet.set_vba_name

    worksheet.insert_button('C2', {
                              :macro   => 'say_hello',
                              :caption => 'Hello'
                            }
                            )

    workbook.add_vba_project(File.join(@regression_output, 'vbaProject02.bin'))

    workbook.close
    compare_for_regression
  end
end
