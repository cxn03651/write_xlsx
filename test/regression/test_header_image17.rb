# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionHeaderImage17 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image17
    @xlsx = 'header_image17.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_header(
      '&L&G&C&G&R&G',
      nil,
      {
        image_left:   'test/regression/images/red.jpg',
        image_center: 'test/regression/images/blue.jpg',
        image_right:  'test/regression/images/red.jpg'
      }
    )

    worksheet.set_footer(
      '&L&G&C&G&R&G',
      nil,
      {
        image_left:   'test/regression/images/blue.jpg',
        image_center: 'test/regression/images/red.jpg',
        image_right:  'test/regression/images/blue.jpg'
      }
    )

    workbook.close
    compare_for_regression(
      [],
      {
        'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup']
      }
    )
  end
end
