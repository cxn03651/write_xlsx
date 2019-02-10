# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image08
    @xlsx = 'image08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('B3',
                           'test/regression/images/grey.png',
                           0, 0, 0.5, 0.5
                           )

    workbook.close
    compare_for_regression
  end
end
