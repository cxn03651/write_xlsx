# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format04
    @xlsx = 'cond_format04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(:num_format => 2,       :dxf_index => 1)
    format2 = workbook.add_format(:num_format => '0.000', :dxf_index => 0)

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    worksheet.conditional_formatting('A1',
                                     {
                                       :type     => 'cell',
                                       :format   => format1,
                                       :criteria => '>',
                                       :value    => 2
                                     }
                                     )

    worksheet.conditional_formatting('A2',
                                     {
                                       :type     => 'cell',
                                       :format   => format2,
                                       :criteria => '<',
                                       :value    => 8
                                     }
                                     )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
