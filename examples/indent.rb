#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

##############################################################################
#
# A simple formatting example using Excel::Writer::XLSX.
#
# This program demonstrates the indentation cell format.
#
# reverse(c), May 2004, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook = WriteXLSX.new('indent.xlsx')

worksheet = workbook.add_worksheet
indent1   = workbook.add_format(indent: 1)
indent2   = workbook.add_format(indent: 2)

worksheet.set_column('A:A', 40)

worksheet.write('A1', "This text is indented 1 level",  indent1)
worksheet.write('A2', "This text is indented 2 levels", indent2)

workbook.close
