# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage13 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image13
    @xlsx = 'header_image13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_header(
                         '&L&G&C&G&R&G',
                         nil,
                         {
                           :image_left   => 'test/regression/images/black_72.jpg',
                           :image_center => 'test/regression/images/black_150.jpg',
                           :image_right  => 'test/regression/images/black_300.jpg'
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
