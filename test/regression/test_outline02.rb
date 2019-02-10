# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionOutline02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_outline02
    @xlsx = 'outline02.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet2  = workbook.add_worksheet('Collapsed Rows')

    # Add a general format
    bold = workbook.add_format(:bold => 1)

    # Create a worksheet with outlined rows. This is the same as the
    # previous example except that the rows are collapsed.

    # The group will be collapsed if $hidden is non-zero.
    # The syntax is: set_row(row, height, XF, hidden, level, collapsed)
    #
    worksheet2.set_row(1, nil, nil, 1, 2)
    worksheet2.set_row(2, nil, nil, 1, 2)
    worksheet2.set_row(3, nil, nil, 1, 2)
    worksheet2.set_row(4, nil, nil, 1, 2)
    worksheet2.set_row(5, nil, nil, 1, 1)

    worksheet2.set_row(6,  nil, nil, 1, 2)
    worksheet2.set_row(7,  nil, nil, 1, 2)
    worksheet2.set_row(8,  nil, nil, 1, 2)
    worksheet2.set_row(9,  nil, nil, 1, 2)
    worksheet2.set_row(10, nil, nil, 1, 1)
    worksheet2.set_row(11, nil, nil, 0, 0, 1)

    # Add a column format for clarity
    worksheet2.set_column('A:A', 20)

    worksheet2.set_selection('A14')

    # Add the data, labels and formulas
    worksheet2.write('A1', 'Region', bold)
    worksheet2.write('A2', 'North')
    worksheet2.write('A3', 'North')
    worksheet2.write('A4', 'North')
    worksheet2.write('A5', 'North')
    worksheet2.write('A6', 'North Total', bold)

    worksheet2.write('B1', 'Sales', bold)
    worksheet2.write('B2', 1000)
    worksheet2.write('B3', 1200)
    worksheet2.write('B4', 900)
    worksheet2.write('B5', 1200)
    worksheet2.write('B6', '=SUBTOTAL(9,B2:B5)', bold, 4300)

    worksheet2.write('A7',  'South')
    worksheet2.write('A8',  'South')
    worksheet2.write('A9',  'South')
    worksheet2.write('A10', 'South')
    worksheet2.write('A11', 'South Total', bold)

    worksheet2.write('B7',  400)
    worksheet2.write('B8',  600)
    worksheet2.write('B9',  500)
    worksheet2.write('B10', 600)
    worksheet2.write('B11', '=SUBTOTAL(9,B7:B10)', bold, 2100)

    worksheet2.write('A12', 'Grand Total',         bold)
    worksheet2.write('B12', '=SUBTOTAL(9,B2:B10)', bold, 6400)

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
