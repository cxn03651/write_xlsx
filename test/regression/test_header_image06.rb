# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_header_image06
    @xlsx = 'header_image06.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.set_header('&L&G', nil, { :image_left   => 'test/regression/images/red.jpg' })
    worksheet2.set_header('&L&G', nil, { :image_left   => 'test/regression/images/blue.jpg' })


    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ],
                                  'xl/worksheets/sheet2.xml' => [ '<pageMargins', '<pageSetup' ],
                                }
                                )
  end
end
