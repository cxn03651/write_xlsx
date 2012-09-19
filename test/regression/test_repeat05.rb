# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRepeat05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_repeat05
    @xlsx = 'repeat05.xlsx'
    workbook   = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.repeat_rows(0)
    worksheet3.repeat_rows(2, 3)
    worksheet3.repeat_columns('B:F')

    worksheet1.write('A1', 'Foo')

    workbook.close
    compare_xlsx_for_regression(
                                File.join(@regression_output, @xlsx),
                                @xlsx,
                                [
                                 'xl/printerSettings/printerSettings1.bin',
                                 'xl/printerSettings/printerSettings2.bin',
                                 'xl/worksheets/_rels/sheet1.xml.rels',
                                 'xl/worksheets/_rels/sheet3.xml.rels'
                                ],
                                {
                                 '[Content_Types].xml'      => ['<Default Extension="bin"'],
                                 'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup'],
                                 'xl/worksheets/sheet3.xml' => ['<pageMargins', '<pageSetup']
                                }
                                )
  end
end
