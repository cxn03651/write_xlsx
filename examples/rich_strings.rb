#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# An Excel::Writer::XLSX example showing how to use "rich strings", i.e.,
# strings with multiple formatting.
#
# reverse(c), February 2011, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('rich_strings.xlsx')
worksheet = workbook.add_worksheet

worksheet.set_column('A:A', 30)

# Set some formats to use.
bold   = workbook.add_format(bold: 1)
italic = workbook.add_format(italic: 1)
red    = workbook.add_format(color: 'red')
blue   = workbook.add_format(color: 'blue')
center = workbook.add_format(align: 'center')
superc = workbook.add_format(font_script: 1)

# Write some strings with multiple formats.
worksheet.write_rich_string(
  'A1',
  'This is ', bold, 'bold', ' and this is ', italic, 'italic'
)

worksheet.write_rich_string(
  'A3',
  'This is ', red, 'red', ' and this is ', blue, 'blue'
)

worksheet.write_rich_string(
  'A5',
  'Some ', bold, 'bold text', ' centered', center
)

worksheet.write_rich_string(
  'A7',
  italic, 'j = k', superc, '(n-1)', center
)

workbook.close
