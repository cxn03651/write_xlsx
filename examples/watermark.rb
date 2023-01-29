#!/usr/bin/env ruby

#######################################################################
#
# An example of adding a worksheet watermark image using the WriteXLSX
# rubygem. This is based on the method of putting an image in the worksheet
# header as suggested in the Microsoft documentation:
# https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('watermark.xlsx')
worksheet = workbook.add_worksheet

# Set a worksheet header with the watermark image.
dirname = File.dirname(File.expand_path(__FILE__))
worksheet.set_header(
  '&C&C&[Picture]', nil,
  { image_center: File.join(dirname, 'watermark.png') }
)

workbook.close
