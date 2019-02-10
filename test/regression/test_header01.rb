# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeader01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header01
    @xlsx = 'header01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.set_header('&L&P', nil, { :scale_with_doc => 0 } )

    workbook.close

    compare_for_regression(
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]
                                })
  end
end
