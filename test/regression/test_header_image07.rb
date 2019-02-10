# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image07
    @xlsx = 'header_image07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('B3', 'test/regression/images/red.jpg' )

    worksheet.set_header('&L&G', nil, { :image_left   => 'test/regression/images/blue.jpg' })


    workbook.close
    compare_for_regression(
                                [],
                                {'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]}
                                )
  end
end
