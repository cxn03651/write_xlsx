# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefinedName01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_defined_name01
    @xlsx = 'defined_name01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet('Sheet 3')

    worksheet1.print_area('A1:E6')
    worksheet1.autofilter('F1:G1')
    worksheet1.write('G1', 'Filter')
    worksheet1.write('F1', 'Auto')
    worksheet1.fit_to_pages(2, 2)

    workbook.define_name("'Sheet 3'!Bar", "='Sheet 3'!$A$1")
    workbook.define_name("Abc",           "=Sheet1!$A$1")
    workbook.define_name("Baz",           "=0.98")
    workbook.define_name("Sheet1!Bar",    "=Sheet1!$A$1")
    workbook.define_name("Sheet2!Bar",    "=Sheet2!$A$1")
    workbook.define_name("Sheet2!aaa",    "=Sheet2!$A$1")
    workbook.define_name("_Egg",          "=Sheet1!$A$1")
    workbook.define_name("_Fog",          "=Sheet1!$A$1")

    workbook.close
    compare_for_regression(
      ["xl/printerSettings/printerSettings1.bin",
       "xl/worksheets/_rels/sheet1.xml.rels"],
      {
        '[Content_Types].xml' => ['<Default Extension="bin"'],
        'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup']
      }
    )
  end
end
