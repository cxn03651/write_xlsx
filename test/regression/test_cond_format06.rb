# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format06
    @xlsx = 'cond_format06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(
                                  :pattern  => 15,
                                  :fg_color => '#FF0000',
                                  :bg_color => '#FFFF00'
                                  )

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    worksheet.conditional_formatting('A1',
                                     {
                                       :type     => 'cell',
                                       :format   => format1,
                                       :criteria => '>',
                                       :value    => 7
                                     }
                                     )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
