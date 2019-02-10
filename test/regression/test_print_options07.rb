# -*- coding: utf-8 -*-
require 'helper'

class TestPrintOptions07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_print_options07
    @xlsx = 'print_options07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write(0, 0, 'Foo')

    worksheet.paper = 9
    worksheet.vertical_dpi = 200

    worksheet.print_black_and_white

    workbook.close
    compare_for_regression(
      [],
      {'xl/worksheets/sheet1.xml' => ['<pageMargins']}
    )
  end
end
