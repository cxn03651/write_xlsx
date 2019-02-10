# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image05
    @xlsx = 'header_image05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_header('&L&G', nil, { :image_left   => 'test/regression/images/red.jpg' })
    worksheet.set_footer('&L&G', nil, { :image_left   => 'test/regression/images/blue.jpg' })


    workbook.close
    compare_for_regression(
                                [],
                                {'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]}
                                )
  end
end
