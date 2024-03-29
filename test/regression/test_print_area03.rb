# -*- coding: utf-8 -*-

require 'helper'

class TestPrintArea03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_print_area03
    @xlsx = 'print_area03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.print_area('A1:XFD1')

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_for_regression(
      [
        'xl/printerSettings/printerSettings1.bin',
        'xl/worksheets/_rels/sheet1.xml.rels'
      ],
      {
        '[Content_Types].xml'      => ['<Default Extension="bin"'],
        'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup']
      }
    )
  end
end
