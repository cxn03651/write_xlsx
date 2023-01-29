# -*- coding: utf-8 -*-

require 'helper'

class TestPrintArea08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_print_area08
    @xlsx = 'print_area08.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.print_area('A1:A1')
    worksheet1.write('A1', 'Foo')

    worksheet2.print_area('A1:A1')
    worksheet2.write('A1', 'Foo')

    workbook.close
    compare_for_regression(
      [],
      {
        'xl/worksheets/sheet1.xml' => ['<pageSetup'],
        'xl/worksheets/sheet2.xml' => ['<pageSetup']
      }
    )
  end
end
