# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable14 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table14
    @xlsx = 'table14.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format1 = workbook.add_format(:num_format => '0.00;[Red]0.00',       :dxf_index => 2)
    format2 = workbook.add_format(:num_format => '0.00_ ;\-0.00\ ',      :dxf_index => 1)
    format3 = workbook.add_format(:num_format => '0.00_ ;[Red]\-0.00\ ', :dxf_index => 0)

    data = [
            [ 'Foo', 1234, 2000, 4321 ],
            [ 'Bar', 1256, 4000, 4320 ],
            [ 'Baz', 2234, 3000, 4332 ],
            [ 'Bop', 1324, 1000, 4333 ]
           ]

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table(
                        'C2:F6',
                        {
                          :data    => data,
                          :columns => [
                                       {},
                                       {:format => format1},
                                       {:format => format2},
                                       {:format => format3}
                                      ]
                        }
                        )

    workbook.close
    compare_for_regression(
                                nil,
                                {  'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
