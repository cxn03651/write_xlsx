# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format02
    @xlsx = 'format02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(0, 30)

    format1 = workbook.add_format(
                                  :font     => "Arial",
                                  :bold     => 1,
                                  :locked   => 1,
                                  :rotation => 0,
                                  :align    => "left",
                                  :valign   => "bottom"
                                  )

    format2 = workbook.add_format(
                                  :font     => "Arial",
                                  :bold     => 1,
                                  :locked   => 1,
                                  :rotation => 90,
                                  :align    => "center",
                                  :valign   => "bottom"
                                  )

    worksheet.write('A1', 'Foo', format1)
    worksheet.write('B1', 'Bar', format2)

    workbook.close
    compare_for_regression(
                                {},
                                {'xl/workbook.xml' => ['<workbookView']}
                                )
  end
end
