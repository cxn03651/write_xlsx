# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionIgnoreError05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_ignore_error05
    @xlsx = 'ignore_error05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string('A1', '123')
    worksheet.write_formula('A2', '=1/0', nil, '#DIV/0!')

    worksheet.ignore_errors(
      number_stored_as_text: 'A1',
      eval_error:            'A2'
    )

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
      {}
    )
  end
end
