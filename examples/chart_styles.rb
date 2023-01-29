#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# An example showing all 48 default chart styles available in Excel 2007
# using Excel::Writer::XLSX.. Note, these styles are not the same as the
# styles available in Excel 2013.
#
# reverse ('(c)'), March 2015, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('chart_styles.xlsx')

# Show the styles for all of these chart types.
chart_types = %w[column area line pie]

chart_types.each do |chart_type|
  # Add a worksheet for each chart type.
  worksheet = workbook.add_worksheet(chart_type.capitalize)
  worksheet.zoom = 30
  style_number = 1

  # Create 48 charts, each with a different style.
  0.step(89, 15) do |row_num|
    0.step(63, 8) do |col_num|
      chart = workbook.add_chart(
        type:     chart_type,
        embedded: 1
      )

      chart.add_series(values: '=Data!$A$1:$A$6')
      chart.set_title(name: "Style #{style_number}")
      chart.set_legend(none: 1)
      chart.set_style(style_number)

      worksheet.insert_chart(row_num, col_num, chart)
      style_number += 1
    end
  end
end

# Create a worksheet with data for the charts.
data = [10, 40, 50, 20, 10, 50]
data_worksheet = workbook.add_worksheet('Data')
data_worksheet.write_col('A1', data)
data_worksheet.hide

workbook.close
