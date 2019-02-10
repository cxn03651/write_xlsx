# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image03
    @xlsx = 'header_image03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_header(
                         '&L&[Picture]&C&G&R&[Picture]',
                         nil,
                         {
                           :image_left   => 'test/regression/images/red.jpg',
                           :image_center => 'test/regression/images/blue.jpg',
                           :image_right  => 'test/regression/images/yellow.jpg'
                         }
                         )

    workbook.close
    compare_for_regression(
                                [],
                                {'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]}
                                )
  end
end
