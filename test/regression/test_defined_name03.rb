# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefinedName03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_defined_name03
    @xlsx = 'defined_name03.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet('sheet One')

    workbook.define_name("Sales", "='sheet One'!G1:H10")

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                ["xl/printerSettings/printerSettings1.bin",
                                 "xl/worksheets/_rels/sheet1.xml.rels"],
                                {
                                  '[Content_Types].xml' => ['<Default Extension="bin"'],
                                  'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup']
                                }
                                )
  end
end
