# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable32 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table32
    @xlsx = 'table32.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the target worksheet.
    worksheet.set_column('C:F', 10.288)

    # Write some strings to order the string table.
    worksheet.write_string('A1', 'Column1')
    worksheet.write_string('B1', 'Column2')
    worksheet.write_string('C1', 'Column3')
    worksheet.write_string('D1', 'Column4')
    worksheet.write_string('E1', 'Total')

    # Add the table.
    worksheet.add_table(
      'C3:F14',
      {
        total_row: 1,
        columns:   [
          { total_string: 'Total' },
          { total_function: 'D5+D9' },
          { total_function: '=SUM([Column3])' },
          { total_function: 'count' }
        ]
      }
    )

    workbook.close
    compare_for_regression(
      [
        'xl/calcChain.xml',
        '[Content_Types].xml',
        'xl/_rels/workbook.xml.rels'
      ],
      {  'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
