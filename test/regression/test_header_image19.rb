# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage19 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image19
    @xlsx = 'header_image19.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/red.jpg')

    worksheet.set_header(
      '&L&G',
      nil,
      {
        :image_left   => 'test/regression/images/red.jpg'
      }
    )

    workbook.close
    compare_for_regression(
      [],
      {
        'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]
      }
    )
  end
end
