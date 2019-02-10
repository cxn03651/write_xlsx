# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionEscapes01 < Test::Unit::TestCase
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

    worksheet.write_formula('A1', %q{=IF(1>2,0,1)},            nil, 1)
    worksheet.write_formula('A2', %q{=CONCATENATE("'","<>&")}, nil, %q{'<>&})
    worksheet.write_formula('A3', %q{=1&"b"},                  nil, %q{1b})
    worksheet.write_formula('A4', %q{="'"},                    nil, %q{'})
    worksheet.write_formula('A5', %q{=""""},                   nil, %q{"})
    worksheet.write_formula('A6', %q{="&" & "&"},              nil, %q{&&})

    worksheet.write_string('A8', %q{"&<>})

    workbook.close
    compare_for_regression(
      [ 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels' ],
      {
        'xl/workbook.xml' => ['<workbookView']
      }
    )
  end
end
