# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionGridlines01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_gridlines01
    @xlsx = 'gridlines01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.hide_gridlines(2)

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_for_regression(
      %w[
         xl/printerSettings/printerSettings1.bin
         xl/worksheets/_rels/sheet1.xml.rels
      ],
      {
        '[Content_Types].xml'      => ['<Default Extension="bin"'],
        'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup']
      }
    )
  end
end
