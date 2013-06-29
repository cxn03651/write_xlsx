# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_utf8_05
    @xlsx = 'utf8_05.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', '="Café"', nil, 'Café')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                [ 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels' ],
                                {'xl/workbook.xml' => ['<workbookView']}
                                )
  end
end
