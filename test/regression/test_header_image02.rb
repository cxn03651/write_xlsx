# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image02
    @xlsx = 'header_image02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_header(
                         '&L&G&C&G',
                         nil,
                         {
                           :image_left   => 'test/regression/images/red.jpg',
                           :image_center => 'test/regression/images/blue.jpg'
                         }
                         )

    workbook.close
    compare_for_regression(
                                [],
                                {'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]}
                                )
  end
end
