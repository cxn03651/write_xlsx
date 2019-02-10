# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionImage19 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_image19
    @xlsx = 'image19.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('C2',
                           File.join(@test_dir, 'regression', 'images/train.jpg'))

    workbook.close
    compare_for_regression
  end
end
