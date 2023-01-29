# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionComment13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment13
    @xlsx = 'comment13.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment(
      'B2',
      'Some text',
      font:        'Courier',
      font_size:   10,
      font_family: 3
    )

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      ['xl/styles.xml'],
      {}
    )
  end
end
