# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_cond_format11
    @xlsx = 'cond_format11.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    format = workbook.add_format(
                                 :bg_color => '#FFFF00',
                                 :fg_color => '#FF0000',
                                 :pattern  => 1
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
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                nil,
                                { 'xl/workbook.xml' => ['<workbookView'] })
  end
end
