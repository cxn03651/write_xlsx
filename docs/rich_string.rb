require 'write_xlsx'

workbook = WriteXLSX.new('rich_string.xlsx')
worksheet = workbook.add_worksheet

bold   = workbook.add_format(bold:        1)
italic = workbook.add_format(italic:      1)
red    = workbook.add_format(color:       'red')
blue   = workbook.add_format(color:       'blue')
center = workbook.add_format(align:       'center')
superscript  = workbook.add_format(font_script: 1)

# Write some strings with multiple formats.
worksheet.write_rich_string('A1',
  'This is ', bold, 'bold', ' and this is ', italic, 'italic')

worksheet.write_rich_string('A3',
  'This is ', red, 'red', ' and this is ', blue, 'blue')

worksheet.write_rich_string('A5',
  'Some ', bold, 'bold text', ' centered', center)

worksheet.write_rich_string('A7',
  italic, 'j = k', superscript, '(n-1)', center)

workbook.close
