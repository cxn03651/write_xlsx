# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionBackground06 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background06
    @xlsx = 'background06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/logo.jpg')
    worksheet.set_background('test/regression/images/logo.jpg')

    worksheet.set_header(
      '&C&G', nil, :image_center => 'test/regression/images/blue.jpg'
    )

    workbook.close
    compare_for_regression(
      [],
      'xl/worksheets/sheet1.xml' => ['<pageSetup']
    )
  end
end
