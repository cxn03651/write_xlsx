# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionOutline03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_outline03
    @xlsx = 'outline03.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet3  = workbook.add_worksheet('Outline Columns')

    # Add a general format
    bold = workbook.add_format(:bold => 1)

    # Example 3: Create a worksheet with outlined columns.
    data = [
            ['Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total'],
            ['North', 50,    20,    15,    25,    65,    80],
            ['South', 10,    20,    30,    50,    50,    50],
            ['East',  45,    75,    50,    15,    75,    100],
            ['West',  15,    15,    55,    35,    20,    50]
           ]
    # Add bold format the first row
    worksheet3.set_row(0, nil, bold)

    # Syntax: set_column(col1, col2, wodth, XF, hidden, level, collapsed)
    worksheet3.set_column('A:A', 10, bold)
    worksheet3.set_column('B:G', 6, nil, 0, 1)
    worksheet3.set_column('H:H', 10)

    # Write the data and a formula
    worksheet3.write_col('A1', data)
    worksheet3.write('H2', '=SUM(B2:G2)', nil, 255)
    worksheet3.write('H3', '=SUM(B3:G3)', nil, 210)
    worksheet3.write('H4', '=SUM(B4:G4)', nil, 360)
    worksheet3.write('H5', '=SUM(B5:G5)', nil, 190)
    worksheet3.write('H6', '=SUM(H2:H5)', bold, 1015)

    workbook.close
    compare_for_regression(
      [
        'xl/calcChain.xml',
        '[Content_Types].xml',
        'xl/_rels/workbook.xml.rels'
      ],
      {
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
