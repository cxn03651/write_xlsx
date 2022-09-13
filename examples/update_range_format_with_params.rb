#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A simple example of how to use the write_xlsx gem to
# update format of the range of cells.
#

require 'write_xlsx'

workbook = WriteXLSX.new('update_range_format_with_params.xlsx')
worksheet = workbook.add_worksheet

common_format = workbook.add_format(align: 'center', border: 1)

table_contents = [
  %w[Table Header Contents],
  %w[table body contents],
  %w[table body contents],
  %w[table body contents]
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
