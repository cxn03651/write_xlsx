#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

##############################################################################
#
# An example of adding document properties to a Excel::Writer::XLSX file.
#
# reverse('©'), August 2008, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('properties.xlsx')
worksheet = workbook.add_worksheet

workbook.set_properties(
  title:    'This is an example spreadsheet',
  subject:  'With document properties',
  author:   'John McNamara',
  manager:  'Dr. Heinz Doofenshmirtz',
  company:  'of Wolves',
  category: 'Example spreadsheets',
  keywords: 'Sample, Example, Properties',
  comments: 'Created with Perl and Excel::Writer::XLSX',
  status:   'Quo'
)

worksheet.set_column('A:A', 70)
worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

workbook.close
