# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table09
    @xlsx = 'table09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the taget worksheet.
    worksheet.set_column('B:K', 10.288)

    # Add some strings to order the string table.
    worksheet.write_string('A1', 'Column1')
    worksheet.write_string('B1', 'Column2')
    worksheet.write_string('C1', 'Column3')
    worksheet.write_string('D1', 'Column4')
    worksheet.write_string('E1', 'Column5')
    worksheet.write_string('F1', 'Column6')
    worksheet.write_string('G1', 'Column7')
    worksheet.write_string('H1', 'Column8')
    worksheet.write_string('I1', 'Column9')
    worksheet.write_string('J1', 'Column10')
    worksheet.write_string('K1', 'Total')

    # Populate the data range.
    data = [0, 0, 0, nil, nil, 0, 0, 0, 0, 0]
    worksheet.write_row('B4', data)
    worksheet.write_row('B5', data)

    # Add the table.
    worksheet.add_table(
                        'B3:K6',
                        {
                          :total_row => 1,
                          :columns   => [
                                         {:total_string => 'Total'},
                                         {},
                                         {:total_function => 'Average'},
                                         {:total_function => 'COUNT'},
                                         {:total_function => 'count_nums'},
                                         {:total_function => 'max'},
                                         {:total_function => 'min'},
                                         {:total_function => 'sum'},
                                         {:total_function => 'std Dev'},
                                         {:total_function => 'var'}
                                        ]
                        }
                        )

    workbook.close
    compare_for_regression(
                                [ 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels' ],
                                {  'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
