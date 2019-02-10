# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionArrayFormula02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_array_formula02
    @xlsx = 'array_formula02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:bold => 1)

    data = [0, 0, 0]

    worksheet.write_col('B1', data)
    worksheet.write_col('C1', data)

    worksheet.write_array_formula('A1:A3', '{=SUM(B1:C1*B2:C2)}', format, 0)

    workbook.close
    compare_for_regression(
      [ 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels' ],
      {'xl/workbook.xml' => ['<workbookView']}
    )
  end
end
