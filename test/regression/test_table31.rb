# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable31 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table31
    @xlsx = 'table31.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(
      bg_color: '#FFFF00',
      fg_color: '#FF0000',
      pattern:  6
    )

    data = [
      ['Foo', 1234, 2000, 4321],
      ['Bar', 1256, 4000, 4320],
      ['Baz', 2234, 3000, 4332],
      ['Bop', 1324, 1000, 4333]
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
          { format: format1 }
        ]
      }
    )

    workbook.close
    compare_for_regression
  end
end
