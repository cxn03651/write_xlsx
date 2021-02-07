#!/usr/bin/env ruby

#######################################################################
#
# Example of how to set Excel worksheet tab colours.
#
# reverse(c), May 2006, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook = Excel::Writer::XLSX.new('tab_colors.xlsx')

worksheet1 = workbook.add_worksheet
worksheet2 = workbook.add_worksheet
worksheet3 = workbook.add_worksheet
worksheet4 = workbook.add_worksheet

# Worksheet1 will have the default tab colour.
worksheet2.tab_color = 'red'
worksheet3.tab_color = 'green'
worksheet4.tab_color = '#FF6600'    # Orange

workbook.close
