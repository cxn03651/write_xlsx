# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRepeat02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_repeat02
    @xlsx = 'repeat02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.repeat_columns('A:A')

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
