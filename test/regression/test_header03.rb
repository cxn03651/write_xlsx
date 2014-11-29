# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeader03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_header03
    @xlsx = 'header03.xlsx'
    workbook    = WriteXLSX.new(@xlsx)
    worksheet   = workbook.add_worksheet

    worksheet.set_footer('&L&P', nil,
                         { :scale_with_doc => 0,
                           :align_with_margins => 0 }
                         )

    workbook.close

    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]
                                })
  end
end
