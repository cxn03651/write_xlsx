# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format03
    @xlsx = 'cond_format03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(:font_strikeout => 1, :dxf_index => 1)
    format2 = workbook.add_format(:underline      => 1, :dxf_index => 0)

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    worksheet.conditional_formatting('A1',
                                     {
                                       :type     => 'cell',
                                       :format   => format1,
                                       :criteria => 'between',
                                       :minimum  => 2,
                                       :maximum => 6
                                     }
                                     )

    worksheet.conditional_formatting('A1',
                                     {
                                       :type     => 'cell',
                                       :format   => format2,
                                       :criteria => 'greater than',
                                       :value    => 1
                                     }
                                     )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
