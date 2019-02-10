# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionPageView01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_page_view01
    @xlsx = 'page_view01.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    worksheet.set_page_view

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
