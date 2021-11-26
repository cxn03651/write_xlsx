#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# An example of setting a worksheet background image with Excel::Writer::XLSX.
#
# Copyright 2000-2021, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('background.xlsx')
worksheet = workbook.add_worksheet

worksheet.set_background('republic.png')

workbook.close
