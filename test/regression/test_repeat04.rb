# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionRepeat04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_repeat04
    @xlsx = 'repeat04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet('Sheet 1')

    worksheet.repeat_rows(0)

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
