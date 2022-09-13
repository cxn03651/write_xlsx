# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionArrayFormula04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_array_formula04
    @xlsx = 'array_formula04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_array_formula(
      'A1:A3',
      '{=SUM(B1:C1*B2:C2)}',
      nil,
      0
    )

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
