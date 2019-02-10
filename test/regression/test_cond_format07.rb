# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionCondFormat07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format07
    @xlsx = 'cond_format07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # We manually set the indices to get the same order as the target file.
    format1 = workbook.add_format(:bg_color => '#FF0000', :dxf_index => 1)
    format2 = workbook.add_format(:bg_color => '#92D050', :dxf_index => 0)

    data = [
            [ 90, 80,  50, 10,  20,  90,  40, 90,  30,  40 ],
            [ 20, 10,  90, 100, 30,  60,  70, 60,  50,  90 ],
            [ 10, 50,  60, 50,  20,  50,  80, 30,  40,  60 ],
            [ 10, 90,  20, 40,  10,  40,  50, 70,  90,  50 ],
            [ 70, 100, 10, 90,  10,  10,  20, 100, 100, 40 ],
            [ 20, 60,  10, 100, 30,  10,  20, 60,  100, 10 ],
            [ 10, 60,  10, 80,  100, 80,  30, 30,  70,  40 ],
            [ 30, 90,  60, 10,  10,  100, 40, 40,  30,  40 ],
            [ 80, 90,  10, 20,  20,  50,  80, 20,  60,  90 ],
            [ 60, 80,  30, 30,  10,  50,  80, 60,  50,  30 ]
           ]

    worksheet.write_col('A1', data)

    worksheet.conditional_formatting('A1:J10',
                                     {
                                       :type     => 'cell',
                                       :format   => format1,
                                       :criteria => '>=',
                                       :value    => 50
                                     }
                                     )

    worksheet.conditional_formatting('A1:J10',
                                     {
                                       :type     => 'cell',
                                       :format   => format2,
                                       :criteria => '<',
                                       :value    => 50
                                     }
                                     )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
