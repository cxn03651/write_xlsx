# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable23 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table23
    @xlsx = 'table23.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the target worksheet.
    worksheet.set_column('B:F', 10.288)

    # Write some strings to order the string table.
    worksheet.write_string('A1', 'Column1')
    worksheet.write_string('F1', 'Total')
    worksheet.write_string('B1', "Column'")
    worksheet.write_string('C1', 'Column#')
    worksheet.write_string('D1', 'Column[')
    worksheet.write_string('E1', 'Column]')

    # Populate the data range.
    # data =  [0, 0, 0, nil, nil, 0, 0, 0, 0, 0]
    # worksheet.write_row('B4', data)
    # worksheet.write_row('B5', data)

    # Add the table.
    worksheet.add_table(
      'B3:F9',
      {
        :total_row => 1,
        :columns   => [
          { :header => 'Column1', :total_string   => 'Total' },
          { :header => "Column'", :total_function => 'sum' },
          { :header => 'Column#', :total_function => 'sum' },
          { :header => 'Column[', :total_function => 'sum' },
          { :header => 'Column]', :total_function => 'sum' }
        ]
      }
    )

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
      {  'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
