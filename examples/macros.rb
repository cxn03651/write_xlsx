#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# An example of adding macros to an WriteXLSX file using VBA project file
# extracted from an existing Excel xlsm file.
#
# The +extract_vba.rb+ utility supplied with WriteXLSX can be used to extract
# the vbaProject.bin file.
#
# An embedded macro is connected to a form button on the worksheet.
#
# reverse(c), November 2012, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Note the file extension should be .xlsm.
workbook  = WriteXLSX.new('macros.xlsm')
worksheet = workbook.add_worksheet

worksheet.set_column('A:A', 30)

# Add the VBA project binary.
workbook.add_vba_project('./vbaProject.bin')

# Show text for the end user.
worksheet.write('A3', 'Press the button to say hello.')

# Add a button tied to a macro in the VBA project.
worksheet.insert_button(
  'B3',
  macro:   'say_hello',
  caption: 'Press Me',
  width:   80,
  height:  30
)

workbook.close
