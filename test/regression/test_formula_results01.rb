# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormulaResults01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_formula_results01
    @xlsx = 'formula_results01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_formula('A1',  '1+1',                 nil, 2)
    worksheet.write_formula('A2',  '"Foo"',               nil, 'Foo')
    worksheet.write_formula('A3',  'IF(B3,FALSE,TRUE)',   nil, 'TRUE')
    worksheet.write_formula('A4',  'IF(B4,TRUE,FALSE)',   nil, 'FALSE')
    worksheet.write_formula('A5',  '#DIV/0!',             nil, '#DIV/0!')
    worksheet.write_formula('A6',  '#N/A',                nil, '#N/A')
    worksheet.write_formula('A7',  '#NAME?',              nil, '#NAME?')
    worksheet.write_formula('A8',  '#NULL!',              nil, '#NULL!')
    worksheet.write_formula('A9',  '#NUM!',               nil, '#NUM!')
    worksheet.write_formula('A10', '#REF!',               nil, '#REF!')
    worksheet.write_formula('A11', '#VALUE!',             nil, '#VALUE!')
    worksheet.write_formula('A12', '1/0',                 nil, '#DIV/0!')

    workbook.close
    compare_for_regression(
      [ 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels' ],
      {}
    )
  end
end
