# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format02
    @xlsx = 'cond_format02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format = nil

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    worksheet.conditional_formatting('A1',
                                     {
                                       :type     => 'cell',
                                       :format   => format,
                                       :criteria => '<',
                                       :value    => 5
                                     }
                                     )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
