#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'write_xlsx'

workbook  = WriteXLSX.new('protection.xlsx')
worksheet = workbook.add_worksheet

# Create some format objects
unlocked = workbook.add_format(locked: 0)
hidden   = workbook.add_format(hidden: 1)

# Format the columns
worksheet.set_column('A:A', 45)
worksheet.set_selection('B3')

# Protect the worksheet
worksheet.protect

# Examples of cell locking and hiding.
worksheet.write('A1', 'Cell B1 is locked. It cannot be edited.')
worksheet.write_formula('B1', '=1+2', nil, 3)    # Locked by default.

worksheet.write('A2', 'Cell B2 is unlocked. It can be edited.')
worksheet.write_formula('B2', '=1+2', unlocked, 3)

worksheet.write('A3', "Cell B3 is hidden. The formula isn't visible.")
worksheet.write_formula('B3', '=1+2', hidden, 3)

worksheet.write('A5', 'Use Menu->Tools->Protection->Unprotect Sheet')
worksheet.write('A6', 'to remove the worksheet protection.')

workbook.close
