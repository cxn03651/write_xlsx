# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionOutline01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_outline01
    @xlsx = 'outline01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet('Outlined Rows')

    # Add a general format
    bold = workbook.add_format(:bold => 1)

    # For outlines the important parameters are hidden and level. Rows with the
    # same level are grouped together. The group will be collapsed if hidden is
    # non-zero. height and XF are assigned default values if they are nil.
    #
    # The syntax is: set_row(row, height, XF, hidden, level, collapsed)
    #
    worksheet1.set_row(1, nil, nil, 0, 2)
    worksheet1.set_row(2, nil, nil, 0, 2)
    worksheet1.set_row(3, nil, nil, 0, 2)
    worksheet1.set_row(4, nil, nil, 0, 2)
    worksheet1.set_row(5, nil, nil, 0, 1)

    worksheet1.set_row(6,  nil, nil, 0, 2)
    worksheet1.set_row(7,  nil, nil, 0, 2)
    worksheet1.set_row(8,  nil, nil, 0, 2)
    worksheet1.set_row(9,  nil, nil, 0, 2)
    worksheet1.set_row(10, nil, nil, 0, 1)

    # Add a column format for clarity
    worksheet1.set_column('A:A', 20)

    # Add the data, labels and formulas
    worksheet1.write('A1', 'Region', bold)
    worksheet1.write('A2', 'North')
    worksheet1.write('A3', 'North')
    worksheet1.write('A4', 'North')
    worksheet1.write('A5', 'North')
    worksheet1.write('A6', 'North Total', bold)

    worksheet1.write('B1', 'Sales', bold)
    worksheet1.write('B2', 1000)
    worksheet1.write('B3', 1200)
    worksheet1.write('B4', 900)
    worksheet1.write('B5', 1200)
    worksheet1.write('B6', '=SUBTOTAL(9,B2:B5)', bold, 4300)

    worksheet1.write('A7',  'South')
    worksheet1.write('A8',  'South')
    worksheet1.write('A9',  'South')
    worksheet1.write('A10', 'South')
    worksheet1.write('A11', 'South Total', bold)

    worksheet1.write('B7',  400)
    worksheet1.write('B8',  600)
    worksheet1.write('B9',  500)
    worksheet1.write('B10', 600)
    worksheet1.write('B11', '=SUBTOTAL(9,B7:B10)', bold, 2100)

    worksheet1.write('A12', 'Grand Total',         bold)
    worksheet1.write('B12', '=SUBTOTAL(9,B2:B10)', bold, 6400)

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
