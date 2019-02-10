# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionOutline04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_outline04
    @xlsx = 'outline04.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet4  = workbook.add_worksheet('Outline levels')

    # Example 4: Show all possible outline levels.
    levels = [
              'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 'Level 6',
              'Level 7', 'Level 6', 'Level 5', 'Level 4', 'Level 3', 'Level 2',
              'Level 1'
             ]

    worksheet4.write_col('A1', levels)

    worksheet4.set_row(0,  nil, nil, nil, 1)
    worksheet4.set_row(1,  nil, nil, nil, 2)
    worksheet4.set_row(2,  nil, nil, nil, 3)
    worksheet4.set_row(3,  nil, nil, nil, 4)
    worksheet4.set_row(4,  nil, nil, nil, 5)
    worksheet4.set_row(5,  nil, nil, nil, 6)
    worksheet4.set_row(6,  nil, nil, nil, 7)
    worksheet4.set_row(7,  nil, nil, nil, 6)
    worksheet4.set_row(8,  nil, nil, nil, 5)
    worksheet4.set_row(9,  nil, nil, nil, 4)
    worksheet4.set_row(10, nil, nil, nil, 3)
    worksheet4.set_row(11, nil, nil, nil, 2)
    worksheet4.set_row(12, nil, nil, nil, 1)

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
