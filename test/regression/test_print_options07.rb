# -*- coding: utf-8 -*-
require 'helper'

class TestPrintOptions07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_print_options07
    @xlsx = 'print_options07.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write(0, 0, 'Foo')

    worksheet.paper = 9
    worksheet.vertical_dpi = 200

    worksheet.print_black_and_white

    workbook.close
    compare_xlsx_for_regression(
      File.join(@regression_output, @xlsx),
      @xlsx,
      [],
      {'xl/worksheets/sheet1.xml' => ['<pageMargins']}
    )
  end
end
