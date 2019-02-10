# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionExcel2003Style06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_excel2003_style06
    @xlsx = 'excel2003_style06.xlsx'
    workbook    = WriteXLSX.new(@io, :excel2003_style => true)
    worksheet   = workbook.add_worksheet

    worksheet.insert_image('B3', 'test/regression/images/red.jpg', 4, 3)

    workbook.close
    compare_for_regression(
                                [],
                                {
                                  'xl/drawings/drawing1.xml' =>
                                  [
                                   '<xdr:cNvPr', '<a:picLocks', '<a:srcRect />', '<xdr:spPr', '<a:noFill />'
                                  ]
                                }
                                )
  end
end
