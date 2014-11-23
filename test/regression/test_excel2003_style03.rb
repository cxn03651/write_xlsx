# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_excel2003_style03
    @xlsx = 'excel2003_style03.xlsx'
    workbook    = WriteXLSX.new(@xlsx, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    worksheet.paper = 9

    worksheet.set_header('Page &P')
    worksheet.set_footer('&A')

    bold = workbook.add_format(:bold => true)

    worksheet.write('A1', 'Foo')
    worksheet.write('A2', 'Bar', bold)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
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
