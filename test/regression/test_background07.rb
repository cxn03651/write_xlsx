# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionBackground07 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background07
    @xlsx = 'background07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.insert_image('E9', 'test/regression/images/logo.jpg')
    worksheet.set_background('test/regression/images/logo.jpg')

    worksheet.set_header(
      '&C&G', nil, :image_center => 'test/regression/images/blue.jpg'
    )

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      [],
      'xl/worksheets/sheet1.xml' => ['<pageSetup']
    )
  end
end
