#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# An example of adding macros to an Excel::Writer::XLSX file using
# a VBA project file extracted from an existing Excel xlsm file.
#
# The C<extract_vba> utility supplied with Excel::Writer::XLSX can be
# used to extract the vbaProject.bin file.
#
# reverse('(c)'), November 2012, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Note the file extension should be .xlsm.
workbook  = WriteXLSX.new('add_vba_project.xlsm')
worksheet = workbook.add_worksheet

worksheet.set_column('A:A', 50)

# Add the VBA project binary.
workbook.add_vba_project(File.join(File.dirname(__FILE__), 'vbaProject.bin'))

# Show text for the end user.
worksheet.write('A1', 'Run the SampleMacro embedded in this file.')
worksheet.write('A2', 'You may have to turn on the Excel Developer option first.')

# Call a user defined function from the VBA project.
worksheet.write('A6', 'Result from a user defined function:')
worksheet.write('B6', '=MyFunction(7)')

workbook.close
