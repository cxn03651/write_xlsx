# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image10
    @xlsx = 'header_image10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.set_header('&L&G', nil, { :image_left   => 'test/regression/images/red.jpg' })

    worksheet2.write('A1', 'Foo')
    worksheet2.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet2.comments_author = 'John'

    workbook.close
    compare_for_regression(
                                [],
                                {
                                  'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ],
                                  'xl/worksheets/sheet2.xml' => [ '<pageMargins', '<pageSetup' ],
                                }
                                )
  end
end
