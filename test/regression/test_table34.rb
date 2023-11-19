# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable34 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table34
    @xlsx = 'table34.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(num_format: "0.0000")

    data = [
      ['Foo', 1234, 0, 4321],
      ['Bar', 1256, 0, 4320],
      ['Baz', 2234, 0, 4332],
      ['Bop', 1324, 0, 4333]
    ]

    # Set the column width to match the target worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table(
      'C2:F6',
      {
        data:    data,
        columns: [
          {},
          {},
          {},
          {
            formula: 'Table1[[#This Row],[Column3]]',
            format:  format
          }
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
