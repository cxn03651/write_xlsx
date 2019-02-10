# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeader03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header03
    @xlsx = 'header03.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.set_footer('&L&P', nil,
                         { :scale_with_doc => 0,
                           :align_with_margins => 0 }
                         )

    workbook.close

    compare_for_regression(
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]
                                })
  end
end
