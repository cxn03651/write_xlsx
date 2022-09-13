#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'write_xlsx'

workbook  = WriteXLSX.new('comments1.xlsx')
worksheet = workbook.add_worksheet

worksheet.write('A1', 'Hello')
worksheet.write_comment('A1', 'This is a comment')

workbook.close
