#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'write_xlsx'

workbook  = WriteXLSX.new('keep_leading_zeros.xlsx')
worksheet = workbook.add_worksheet

worksheet.keep_leading_zeros(true)
worksheet.write('A1', '001')
worksheet.write('B1', 'written as string.')
worksheet.write('A2', '012')
worksheet.write('B2', 'written as string.')
worksheet.write('A3', '123')
worksheet.write('B3', 'written as number.')

workbook.close
