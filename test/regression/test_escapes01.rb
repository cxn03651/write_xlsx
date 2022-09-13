# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionEscapes01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_escapes01
    @xlsx = 'escapes01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet('5&4')

    worksheet.write_formula('A1', '=IF(1>2,0,1)',            nil, 1)
    worksheet.write_formula('A2', %q{=CONCATENATE("'","<>&")}, nil, "'<>&")
    worksheet.write_formula('A3', '=1&"b"',                  nil, '1b')
    worksheet.write_formula('A4', %q(="'"),                    nil, "'")
    worksheet.write_formula('A5', '=""""',                   nil, '"')
    worksheet.write_formula('A6', '="&" & "&"',              nil, '&&')

    worksheet.write_string('A8', '"&<>')

    workbook.close
    compare_for_regression(
      ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'],
      {
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
