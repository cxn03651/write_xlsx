# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_05
    @xlsx = 'utf8_05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', '="Café"', nil, 'Café')

    workbook.close
    compare_for_regression(
                                [ 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels' ],
                                {'xl/workbook.xml' => ['<workbookView']}
                                )
  end
end
