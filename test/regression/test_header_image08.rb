# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHeaderImage08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_header_image08
    @xlsx = 'header_image08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    worksheet.set_header('&L&G', nil, { :image_left   => 'test/regression/images/red.jpg' })


    workbook.close
    compare_for_regression(
                                [],
                                {'xl/worksheets/sheet1.xml' => [ '<pageMargins', '<pageSetup' ]}
                                )
  end
end
