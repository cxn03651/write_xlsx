# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage13 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_header_image13
    @xlsx = 'header_image13.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
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
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]
                                }
                                )
  end
end
