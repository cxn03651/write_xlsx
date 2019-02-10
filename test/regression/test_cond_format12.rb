# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format12
    @xlsx = 'cond_format12.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format = workbook.add_format(
                                 :bg_color => '#FFFF00',
                                 :fg_color => '#FF0000',
                                 :pattern  => 12
                                 )

    worksheet.write('A1', 'Hello', format)

    worksheet.write('B3', 10)
    worksheet.write('B4', 20)
    worksheet.write('B5', 30)
    worksheet.write('B6', 40)

    worksheet.conditional_formatting('B3:B6',
                                     {
                                       :type     => 'cell',
                                       :format   => format,
                                       :criteria => 'greater than',
                                       :value    => 20
                                     }
                                     )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
