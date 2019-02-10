# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image06
    @xlsx = 'header_image06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.set_header('&L&G', nil, { :image_left   => 'test/regression/images/red.jpg' })
    worksheet2.set_header('&L&G', nil, { :image_left   => 'test/regression/images/blue.jpg' })


    workbook.close
    compare_for_regression(
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ],
                                  'xl/worksheets/sheet2.xml' => [ '<pageMargins', '<pageSetup' ],
                                }
                                )
  end
end
