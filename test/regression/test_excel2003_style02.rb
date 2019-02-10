# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_excel2003_style02
    @xlsx = 'excel2003_style02.xlsx'
    workbook    = WriteXLSX.new(@io, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    worksheet.paper = 9

    bold = workbook.add_format(:bold => true)

    worksheet.write('A1', 'Foo')
    worksheet.write('A2', 'Bar', bold)

    workbook.close
    compare_for_regression(
                                [
                                 'xl/printerSettings/printerSettings1.bin',
                                 'xl/worksheets/_rels/sheet1.xml.rels'
                                ],
                                {
                                  '[Content_Types].xml'      => ['<Default Extension="bin"'],
                                  'xl/worksheets/sheet1.xml' => ['<pageMargins']
                                }
                                )
  end
end
