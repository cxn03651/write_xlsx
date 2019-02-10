# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeader02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header02
    @xlsx = 'header02.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.set_header('&L&P', nil, { :align_with_margins => 0 } )

    workbook.close

    compare_for_regression(
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]
                                })
  end
end
