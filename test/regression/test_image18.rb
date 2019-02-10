# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage18 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image18
    @xlsx = 'image18.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(1, 96)
    worksheet.set_column('C:C', 18)

    worksheet.insert_image('C2',
                           'test/regression/images/issue32.png', 5, 5)

    workbook.close
    compare_for_regression
  end
end
