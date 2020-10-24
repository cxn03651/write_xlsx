# -*- coding: utf-8 -*-
require 'helper'

class TestUpdateRangeFormatWithParams < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_update_range_format_with_params
    @xlsx = 'update_range_format_with_params.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    common_format = workbook.add_format(align: 'center', border: 1)

    table_contents = [
      ['Table', 'Header', 'Contents'],
      ['table', 'body', 'contents'],
      ['table', 'body', 'contents'],
      ['table', 'body', 'contents']
    ]

    worksheet.write_col(0, 0, table_contents, common_format)
    worksheet.update_range_format_with_params(
      0, 0, 0, 2,
      bold: 1, top: 2, bottom: 2, bg_color: 31
    )
    worksheet.update_range_format_with_params(0, 0, 3, 0, left: 2)
    worksheet.update_range_format_with_params(0, 2, 3, 2, right: 2)
    worksheet.update_range_format_with_params(3, 0, 3, 2, bottom: 2)

    workbook.close
    compare_for_regression(
      nil,
      {'xl/workbook.xml' => ['<workbookView']}
    )
  end
end
