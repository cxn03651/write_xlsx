# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage14 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_header_image14
    @xlsx = 'header_image14.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.set_header(
                         '&L&G&C&G&R&G',
                         nil,
                         {
                           :image_left   => 'test/regression/images/black_72e.png',
                           :image_center => 'test/regression/images/black_150e.png',
                           :image_right  => 'test/regression/images/black_300e.png'
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
