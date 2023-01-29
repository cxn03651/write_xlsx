# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionIgnoreError06 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_ignore_error06
    @xlsx = 'ignore_error06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_formula('A1', '=B1')
    worksheet.write_formula('A2', '=B1')
    worksheet.write_formula('A3', '=B3')

    worksheet.ignore_errors(
      formula_differs: 'A2'
    )

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
      {}
    )
  end
end
