# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx'

class TestExampleMatch < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close
  end

  def test_multi_line
    @xlsx = 'multi_line.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write(0, 0, "Hi Excel!\n1234\nHi, again!")

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_a_simple
    @xlsx = 'a_simple.xlsx'
    # Create a new workbook called simple.xls and add a worksheet
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # The general syntax is write(row, column, token). Note that row and
    # column are zero indexed
    #

    # Write some text
    worksheet.write(0, 0, "Hi Excel!")

    # Write some numbers
    worksheet.write(2, 0, 3)          # Writes 3
    worksheet.write(3, 0, 3.00000)    # Writes 3
    worksheet.write(4, 0, 3.00001)    # Writes 3.00001
    worksheet.write(5, 0, 3.14159)    # TeX revision no.?

    # Write some formulas
    worksheet.write(7, 0, '=A3 + A6')
    worksheet.write(8, 0, '=IF(A5>3,"Yes", "No")')

    # Write a hyperlink
    hyperlink_format = workbook.add_format(
      :color     => 'blue',
      :underline => 1
    )

    worksheet.write(10, 0, 'http://www.ruby-lang.org/', hyperlink_format)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_array_formula
    @xlsx = 'array_formula.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Write some test data.
    worksheet.write('B1', [[500, 10], [300, 15]])
    worksheet.write('B5', [[1, 2, 3], [20234, 21003, 10000]])

    # Write an array formula that returns a single value
    worksheet.write('A1', '{=SUM(B1:C1*B2:C2)}')

    # Same as above but more verbose.
    worksheet.write_array_formula('A2:A2', '{=SUM(B1:C1*B2:C2)}')

    # Write an array formula that returns a range of values
    worksheet.write_array_formula('A5:A7', '{=TREND(C5:C7,B5:B7)}')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_autofilter
    @xlsx = 'autofilter.xlsx'
    workbook = WriteXLSX.new(@io)

    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet

    bold = workbook.add_format(:bold => 1)

    # Extract the data embedded at the end of this file.
    data_array = autofilter_data.split("\n")
    headings = data_array.shift.split
    data = []
    data_array.each { |line| data << line.split }

    # Set up several sheets with the same data.
    workbook.worksheets.each do |worksheet|
      worksheet.set_column('A:D', 12)
      worksheet.set_row(0, 20, bold)
      worksheet.write('A1', headings)
    end

    ###############################################################################
    #
    # Example 1. Autofilter without conditions.
    #

    worksheet1.autofilter('A1:D51')
    worksheet1.write('A2', [data])

    ###############################################################################
    #
    #
    # Example 2. Autofilter with a filter condition in the first column.
    #

    # The range in this example is the same as above but in row-column notation.
    worksheet2.autofilter(0, 0, 50, 3)

    # The placeholder "Region" in the filter is ignored and can be any string
    # that adds clarity to the expression.
    #
    worksheet2.filter_column(0, 'Region eq East')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
      region = row_data[0]

      worksheet2.set_row(row, nil, nil, 1) unless region == 'East'
      worksheet2.write(row, 0, row_data)
      row += 1
    end

    ###############################################################################
    #
    #
    # Example 3. Autofilter with a dual filter condition in one of the columns.
    #

    worksheet3.autofilter('A1:D51')

    worksheet3.filter_column('A', 'x eq East or x eq South')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
      region = row_data[0]

      worksheet3.set_row(row, nil, nil, 1) unless %w[East South].include?(region)
      worksheet3.write(row, 0, row_data)
      row += 1
    end

    ###############################################################################
    #
    #
    # Example 4. Autofilter with filter conditions in two columns.
    #

    worksheet4.autofilter('A1:D51')

    worksheet4.filter_column('A', 'x eq East')
    worksheet4.filter_column('C', 'x > 3000 and x < 8000')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
      region = row_data[0]
      volume = row_data[2]

      unless region == 'East' && volume.to_i > 3000 && volume.to_i < 8000
        # Hide row.
        worksheet4.set_row(row, nil, nil, 1)
      end

      worksheet4.write(row, 0, row_data)
      row += 1
    end

    ###############################################################################
    #
    #
    # Example 5. Autofilter with filter for blanks.
    #

    # Create a blank cell in our test data.
    data[5][0] = ''

    worksheet5.autofilter('A1:D51')
    worksheet5.filter_column('A', 'x eq Blanks')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
      region = row_data[0]

      worksheet5.set_row(row, nil, nil, 1) unless region == ''

      worksheet5.write(row, 0, row_data)
      row += 1
    end

    ###############################################################################
    #
    #
    # Example 6. Autofilter with filter for non-blanks.
    #

    worksheet6.autofilter('A1:D51')
    worksheet6.filter_column('A', 'x eq NonBlanks')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
      region = row_data[0]

      worksheet6.set_row(row, nil, nil, 1) unless region != ''

      worksheet6.write(row, 0, row_data)
      row += 1
    end

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def teset_background
    @xlsx     = 'background.xlsx'
    workbook  = WriteXLSX.new(@io)

    worksheet = workbook.add_worksheet
    worksheet.set_background(File.join(@test_dir, 'republic.png'))

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_data_labels
    @xlsx = 'chart_data_labels.xlsx'
    workbook  = WriteXLSX.new(@io)

    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = %w[Number Data Text]
    data = [
      [2,  3,  4,  5,  6,  7],
      [20, 10, 20, 30, 40, 30],
      %w[Jan Feb Mar Apr May Jun]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    #######################################################################
    #
    # Example with standard data labels.
    #

    # Create a Column chart.
    chart1 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the data series and add the data labels.
    chart1.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1 }
    )

    # Add a chart title.
    chart1.set_title(:name => 'Chart with standard data labels')

    # Turn off the chart legend.
    chart1.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart1, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with value and category data labels.
    #

    # Create a Column chart.
    chart2 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the data series and add the data labels.
    chart2.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1, :category => 1 }
    )

    # Add a chart title.
    chart2.set_title(:name => 'Category and Value data labels')

    # Turn off the chart legend.
    chart2.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D18', chart2, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with standard data labels with different font.
    #

    # Create a Column chart.
    chart3 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the data series and add the data labels.
    chart3.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1,
                        :font  => { :bold     => 1,
                                    :color    => 'red',
                                    :rotation => -30 } }
    )

    # Add a chart title.
    chart3.set_title(:name => 'Data labels with user defined font')

    # Turn off the chart legend.
    chart3.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D34', chart3, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with custom string data labels.
    #

    # Create a Column chart.
    chart4 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the data series and add the data labels.
    chart4.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => {
        :value  => 1,
        :border => { :color => 'red' },
        :fill   => { :color => 'yellow' }
      }
    )

    # Add a chart title.
    chart4.set_title(:name => 'Data labels with formatting')

    # Turn off the chart legend.
    chart4.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D50', chart4, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with custom string data labels.
    #

    # Create a Column chart.
    chart5 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Some custom labels.
    custom_labels = [
      { :value => 'Amy' },
      { :value => 'Bea' },
      { :value => 'Eva' },
      { :value => 'Fay' },
      { :value => 'Liv' },
      { :value => 'Una' }
    ]

    # Configure the data series and add the data labels.
    chart5.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1, :custom => custom_labels }
    )

    # Add a chart title.
    chart5.set_title(:name => 'Chart with custom string data labels')

    # Turn off the chart legend.
    chart5.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D66', chart5, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with custom data labels from cells.
    #

    # Create a Column chart.
    chart6 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Some custom labels.
    custom_labels = [
      { :value => '=Sheet1!$C$2' },
      { :value => '=Sheet1!$C$3' },
      { :value => '=Sheet1!$C$4' },
      { :value => '=Sheet1!$C$5' },
      { :value => '=Sheet1!$C$6' },
      { :value => '=Sheet1!$C$7' }
    ]

    # Configure the data series and add the data labels.
    chart6.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1, :custom => custom_labels }
    )

    # Add a chart title.
    chart6.set_title(:name => 'Chart with custom data labels from cells')

    # Turn off the chart legend.
    chart6.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D82', chart6, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with custom and default data labels.
    #

    # Create a Column chart.
    chart7 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Some custom labels. The nil items will get the default value.
    # We also set a font for the custom items as an extra example.
    custom_labels = [
      { :value => '=Sheet1!$C$2', :font => { :color => 'red' } },
      nil,
      { :value => '=Sheet1!$C$4', :font => { :color => 'red' } },
      { :value => '=Sheet1!$C$5', :font => { :color => 'red' } }
    ]

    # Configure the data series and add the data labels.
    chart7.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1, :custom => custom_labels }
    )

    # Add a chart title.
    chart7.set_title(:name => 'Mixed custom and default data labels')

    # Turn off the chart legend.
    chart7.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D98', chart7, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with deleted custom data labels.
    #

    # Create a Column chart.
    chart8 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Some deleted custom labels and defaults (nil). This allows us to
    # highlight certain values such as the minimum and maximum.
    custom_labels = [
      { :delete => 1 },
      nil,
      { :delete => 1 },
      { :delete => 1 },
      nil,
      { :delete => 1 }
    ]

    # Configure the data series and add the data labels.
    chart8.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1, :custom => custom_labels }
    )

    # Add a chart title.
    chart8.set_title(:name => 'Chart with deleted data labels')

    # Turn off the chart legend.
    chart8.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D114', chart8, { :x_offset => 25, :y_offset => 10 })

    #######################################################################
    #
    # Example with custom string data labels and formatting.
    #

    # Create a Column chart.
    chart9 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Some custom labels.
    custom_labels = [
      { :value => 'Amy', :border => { :color => 'blue' } },
      { :value => 'Bea' },
      { :value => 'Eva' },
      { :value => 'Fay' },
      { :value => 'Liv' },
      { :value => 'Una', :fill => { :color => 'green' } }
    ]

    # Configure the data series and add the data labels.
    chart9.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => {
        :value  => 1,
        :custom => custom_labels,
        :border => { :color => 'red' },
        :fill   => { :color => 'yellow' }
      }
    )

    # Add a chart title.
    chart9.set_title(:name => 'Chart with custom labels and formatting')

    # Turn off the chart legend.
    chart9.set_legend(:none => 1)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D130', chart9, { :x_offset => 25, :y_offset => 10 })

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_gauge
    @xlsx = 'chart_gauge.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    chart_doughnut = workbook.add_chart(:type => 'doughnut', :embedded => 1)
    chart_pie      = workbook.add_chart(:type => 'pie', :embedded => 1)

    # Add some data for the Doughnut and Pie charts. This is set up so the
    # gauge goes from 0-100. It is initially set at 75%.
    worksheet.write_col('H2', ['Donut', 25, 50, 25, 100])
    worksheet.write_col('I2', ['Pie', 75, 1, '=200-I4-I3'])

    # Configure the doughnut chart as the background for the gauge.
    chart_doughnut.add_series(
      :name   => '=Sheet1!$H$2',
      :values => '=Sheet1!$H$3:$H$6',
      :points => [
        { :fill => { :color => 'green' } },
        { :fill => { :color => 'yellow' } },
        { :fill => { :color => 'red' } },
        { :fill => { :none  => 1 } }
      ]
    )

    # Rotate chart so the gauge parts are above the horizontal.
    chart_doughnut.set_rotation(270)

    # Turn off the chart legend.
    chart_doughnut.set_legend(:none => 1)

    # Turn off the chart fill and border.
    chart_doughnut.set_chartarea(
      :border => { :none  => 1 },
      :fill   => { :none  => 1 }
    )

    # Configure the pie chart as the needle for the gauge.
    chart_pie.add_series(
      :name   => '=Sheet1!$I$2',
      :values => '=Sheet1!$I$3:$I$6',
      :points => [
        { :fill => { :none  => 1 } },
        { :fill => { :color => 'black' } },
        { :fill => { :none  => 1 } }
      ]
    )

    # Rotate the pie chart/needle to align with the doughnut/gauge.
    chart_pie.set_rotation(270)

    # Combine the pie and doughnut charts.
    chart_doughnut.combine(chart_pie)

    # Insert the chart into the worksheet.
    worksheet.insert_chart('A1', chart_doughnut)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_scatter06
    @xlsx = 'chart_scatter06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'scatter', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [57708544, 44297600])

    data = [
      [1, 2, 3, 4,  5],
      [2, 4, 6, 8,  10],
      [3, 6, 9, 12, 15]

    ]

    worksheet.write('A1', data)

    chart.add_series(
      :categories => '=Sheet1!$A$1:$A$5',
      :values     => '=Sheet1!$B$1:$B$5'
    )

    chart.add_series(
      :categories => '=Sheet1!$A$1:$A$5',
      :values     => '=Sheet1!$C$1:$C$5'
    )

    chart.set_x_axis(:minor_unit => 1, :major_unit => 3)
    chart.set_y_axis(:minor_unit => 2, :major_unit => 4)

    worksheet.insert_chart('E9', chart)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def autofilter_data
    <<EOS
    Region    Item      Volume    Month
    East      Apple     9000      July
    East      Apple     5000      July
    South     Orange    9000      September
    North     Apple     2000      November
    West      Apple     9000      November
    South     Pear      7000      October
    North     Pear      9000      August
    West      Orange    1000      December
    West      Grape     1000      November
    South     Pear      10000     April
    West      Grape     6000      January
    South     Orange    3000      May
    North     Apple     3000      December
    South     Apple     7000      February
    West      Grape     1000      December
    East      Grape     8000      February
    South     Grape     10000     June
    West      Pear      7000      December
    South     Apple     2000      October
    East      Grape     7000      December
    North     Grape     6000      April
    East      Pear      8000      February
    North     Apple     7000      August
    North     Orange    7000      July
    North     Apple     6000      June
    South     Grape     8000      September
    West      Apple     3000      October
    South     Orange    10000     November
    West      Grape     4000      July
    North     Orange    5000      August
    East      Orange    1000      November
    East      Orange    4000      October
    North     Grape     5000      August
    East      Apple     1000      December
    South     Apple     10000     March
    East      Grape     7000      October
    West      Grape     1000      September
    East      Grape     10000     October
    South     Orange    8000      March
    North     Apple     4000      July
    South     Orange    5000      July
    West      Apple     4000      June
    East      Apple     5000      April
    North     Pear      3000      August
    East      Grape     9000      November
    North     Orange    8000      October
    East      Apple     10000     June
    South     Pear      1000      December
    North     Grape     10000     July
    East      Grape     6000      February
EOS
  end

  def test_chart_area
    @xlsx = 'chart_area.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [40, 40, 50, 30, 25, 50],
      [30, 25, 30, 10,  5, 10]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'area', :embedded => 1)

    # Configure the first series.
    chart.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'Results of sample analysis')
    chart.set_x_axis(:name => 'Test number')
    chart.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart.set_style(11)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_bar
    @xlsx = 'chart_bar.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'bar', :embedded => 1)

    # Configure the first series.
    chart.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'Results of sample analysis')
    chart.set_x_axis(:name => 'Test number')
    chart.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart.set_style(11)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_column
    @xlsx = 'chart_column.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the first series.
    chart.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'Results of sample analysis')
    chart.set_x_axis(:name => 'Test number')
    chart.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart.set_style(11)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_doughnut
    @xlsx = 'chart_doughnut.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = %w[Category Values]
    data = [
      %w[Glazed Chocolate Cream],
      [50,       35,          15]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart1 = workbook.add_chart(:type => 'doughnut', :embedded => 1)

    # Configure the series. Note the use of the array ref to define ranges:
    # [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
    # See below for an alternative syntax.
    chart1.add_series(
      :name       => 'Doughnut sales data',
      :categories => ['Sheet1', 1, 3, 0, 0],
      :values     => ['Sheet1', 1, 3, 1, 1]
    )

    # Add a title.
    chart1.set_title(:name => 'Popular Doughnut Types')

    # Set an Excel chart style. Colors with white outline and shadow.
    chart1.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C2', chart1, 25, 10)

    #
    # Create a Doughnut chart with user defined segment colors.
    #

    # Create an example Doughnut chart like above.
    chart2 = workbook.add_chart(:type => 'doughnut', :embedded => 1)

    # Configure the series and add user defined segment colours.
    chart2.add_series(
      :name       => 'Doughnut sales data',
      :categories => '=Sheet1!$A$2:$A$4',
      :values     => '=Sheet1!$B$2:$B$4',
      :points     => [
        { :fill => { :color => '#FA58D0' } },
        { :fill => { :color => '#61210B' } },
        { :fill => { :color => '#F5F6CE' } }
      ]
    )

    # Add a title.
    chart2.set_title(:name => 'Doughnut Chart with user defined colors')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C18', chart2, 25, 10)

    #
    # Create a Doughnut chart with rotation of the segments.
    #

    # Create an example Doughnut chart like above.
    chart3 = workbook.add_chart(:type => 'doughnut', :embedded => 1)

    # Configure the series.
    chart3.add_series(
      :name       => 'Doughnut sales data',
      :categories => '=Sheet1!$A$2:$A$4',
      :values     => '=Sheet1!$B$2:$B$4'
    )

    # Add a title.
    chart3.set_title(:name => 'Doughnut Chart with segment rotation')

    # Change the angle/rotation of the first segment.
    chart3.set_rotation(90)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C34', chart3, 25, 10)

    #
    # Create a Doughnut chart with user defined hole size.
    #

    # Create an example Doughnut chart like above.
    chart4 = workbook.add_chart(:type => 'doughnut', :embedded => 1)

    # Configure the series.
    chart4.add_series(
      :name       => 'Doughnut sales data',
      :categories => '=Sheet1!$A$2:$A$4',
      :values     => '=Sheet1!$B$2:$B$4'
    )

    # Add a title.
    chart4.set_title(:name => 'Doughnut Chart with user defined hole size')

    # Change the hole size.
    chart4.set_hole_size(33)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C50', chart4, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_line
    @xlsx = 'chart_line.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the first series.
    chart.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'Results of sample analysis')
    chart.set_x_axis(:name => 'Test number')
    chart.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart, 25, 10)

    #
    # Create a stacked chart sub-type
    #
    chart2 = workbook.add_chart(
      :type     => 'line',
      :embedded => 1,
      :subtype  => 'stacked'
    )

    # Configure the first series.
    chart2.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series.
    chart2.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart2.set_title(:name  => 'Stacked Chart')
    chart2.set_x_axis(:name => 'Test number')
    chart2.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart2.set_style(12)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart(
      'D18', chart2,
      { :x_offset => 25, :y_offset => 10 }
    )

    #
    # Create a percent stacked chart sub-type
    #
    chart3 = workbook.add_chart(
      :type     => 'line',
      :embedded => 1,
      :subtype  => 'percent_stacked'
    )

    # Configure the first series.
    chart3.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series.
    chart3.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart3.set_title(:name  => 'Percent Stacked Chart')
    chart3.set_x_axis(:name => 'Test number')
    chart3.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart3.set_style(13)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart(
      'D34', chart3,
      { :x_offset => 25, :y_offset => 10 }
    )

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_pie
    @xlsx = 'chart_pie.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = %w[Category Values]
    data = [
      %w[Apple Cherry Pecan],
      [60,       30,       10]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart1 = workbook.add_chart(:type => 'pie', :embedded => 1)

    # Configure the series. Note the use of the array ref to define ranges:
    # [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
    # See below for an alternative syntax.
    chart1.add_series(
      :name       => 'Pie sales data',
      :categories => ['Sheet1', 1, 3, 0, 0],
      :values     => ['Sheet1', 1, 3, 1, 1]
    )

    # Add a title.
    chart1.set_title(:name => 'Popular Pie Types')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart1.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C2', chart1, 25, 10)

    #
    # Create a Pie chart with user defined segment colors.
    #

    # Create an example Pie chart like above.
    chart2 = workbook.add_chart(:type => 'pie', :embedded => 1)

    # Configure the series and add user defined segment colours.
    chart2.add_series(
      :name       => 'Pie sales data',
      :categories => '=Sheet1!$A$2:$A$4',
      :values     => '=Sheet1!$B$2:$B$4',
      :points     => [
        { :fill => { :color => '#5ABA10' } },
        { :fill => { :color => '#FE110E' } },
        { :fill => { :color => '#CA5C05' } }
      ]
    )

    # Add a title.
    chart2.set_title(:name => 'Pie Chart with user defined colors')

    worksheet.insert_chart('C18', chart2, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_radar
    @xlsx = 'chart_radar.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [30, 60, 70, 50, 40, 30],
      [25, 40, 50, 30, 50, 40]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart1 = workbook.add_chart(:type => 'radar', :embedded => 1)

    # Configure the first series.
    chart1.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart1.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart1.set_title(:name => 'Results of sample analysis')
    chart1.set_x_axis(:name => 'Test number')
    chart1.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart1.set_style(11)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart1, 25, 10)

    #
    # Create a with_markers chart sub-type
    #
    chart2 = workbook.add_chart(
      :type     => 'radar',
      :embedded => 1,
      :subtype  => 'with_markers'
    )

    # Configure the first series.
    chart2.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series.
    chart2.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart2.set_title(:name  => 'Stacked Chart')
    chart2.set_x_axis(:name => 'Test number')
    chart2.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart2.set_style(12)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D18', chart2, 25, 11)

    #
    # Create a filled chart sub-type
    #
    chart3 = workbook.add_chart(
      :type     => 'radar',
      :embedded => 1,
      :subtype  => 'filled'
    )

    # Configure the first series.
    chart3.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series.
    chart3.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart3.set_title(:name  => 'Percent Stacked Chart')
    chart3.set_x_axis(:name => 'Test number')
    chart3.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart3.set_style(13)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D34', chart3, 25, 11)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_scatter
    @xlsx = 'chart_scatter.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'scatter', :embedded => 1)

    # Configure the first series.
    chart.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ $sheetname, $row_start, $row_end, $col_start, $col_end ].$chart->add_series(
    chart.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'Results of sample analysis')
    chart.set_x_axis(:name => 'Test number')
    chart.set_y_axis(:name => 'Sample length (mm)')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_stock
    @xlsx = 'chart_stock.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet
    bold        = workbook.add_format(:bold => 1)
    date_format = workbook.add_format(:num_format => 'dd/mm/yyyy')
    chart       = workbook.add_chart(:type => 'stock', :embedded => 1)

    # Add the worksheet data that the charts will refer to.
    headings = %w[Date High Low Close]
    data = [
      %w[2007-01-01T 2007-01-02T 2007-01-03T 2007-01-04T 2007-01-05T],
      [27.2,  25.03, 19.05, 20.34, 18.5],
      [23.49, 19.55, 15.12, 17.84, 16.34],
      [25.45, 23.05, 17.32, 20.45, 17.34]
    ]

    worksheet.write('A1', headings, bold)

    5.times do |row|
      worksheet.write_date_time(row + 1, 0, data[0][row], date_format)
      worksheet.write(row + 1, 1, data[1][row])
      worksheet.write(row + 1, 2, data[2][row])
      worksheet.write(row + 1, 3, data[3][row])
    end

    worksheet.set_column('A:D', 11)

    # Add a series for each of the High-Low-Close columns.
    chart.add_series(
      :categories => '=Sheet1!$A$2:$A$6',
      :values     => '=Sheet1!$B$2:$B$6'
    )

    chart.add_series(
      :categories => '=Sheet1!$A$2:$A$6',
      :values     => '=Sheet1!$C$2:$C$6'
    )

    chart.add_series(
      :categories => '=Sheet1!$A$2:$A$6',
      :values     => '=Sheet1!$D$2:$D$6'
    )

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'High-Low-Close')
    chart.set_x_axis(:name => 'Date')
    chart.set_y_axis(:name => 'Share price')

    worksheet.insert_chart('E9', chart)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_secondary_axis
    @xlsx = 'chart_secondary_axis.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = %w[Aliens Humans]
    data = [
      [2,  3,  4,  5,  6,  7],
      [10, 40, 50, 20, 10, 50]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the first series.
    chart.add_series(
      :name    => '=Sheet1!$A$1',
      :values  => '=Sheet1!$A$2:$A$7',
      :y2_axis => 1
    )

    chart.add_series(
      :name   => '=Sheet1!$B$1',
      :values => '=Sheet1!$B$2:$B$7'
    )

    chart.set_legend(:position => 'right')

    # Add a chart title and some axis labels.
    chart.set_title(:name => 'Survey results')
    chart.set_x_axis(:name => 'Days')
    chart.set_y_axis(:name => 'Population', :major_gridlines => { :visible => 0 })
    chart.set_y2_axis(:name => 'Laser wounds')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_comments1
    @xlsx = 'comments1.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Hello')
    worksheet.write_comment('A1', 'This is a comment')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_comments2
    @xlsx = 'comments2.xlsx'
    workbook  = WriteXLSX.new(@io)

    text_wrap  = workbook.add_format(:text_wrap => 1, :valign => 'top')
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet
    worksheet7 = workbook.add_worksheet
    worksheet8 = workbook.add_worksheet

    # Variables that we will use in each example.
    cell_text = ''
    comment   = ''

    ###############################################################################
    #
    # Example 1. Demonstrates a simple cell comments without formatting.
    #            comments.
    #

    # Set up some formatting.
    worksheet1.set_column('C:C', 25)
    worksheet1.set_row(2, 50)
    worksheet1.set_row(5, 50)

    # Simple ascii string.
    cell_text = 'Hold the mouse over this cell to see the comment.'

    comment = 'This is a comment.'

    worksheet1.write('C3', cell_text, text_wrap)
    worksheet1.write_comment('C3', comment)

    cell_text = 'This is a UTF-8 string.'
    comment   = '☺'

    worksheet1.write('C6', cell_text, text_wrap)
    worksheet1.write_comment('C6', comment)

    ###############################################################################
    #
    # Example 2. Demonstrates visible and hidden comments.
    #

    # Set up some formatting.
    worksheet2.set_column('C:C', 25)
    worksheet2.set_row(2, 50)
    worksheet2.set_row(5, 50)

    cell_text = 'This cell comment is visible.'

    comment = 'Hello.'

    worksheet2.write('C3', cell_text, text_wrap)
    worksheet2.write_comment('C3', comment, :visible => 1)

    cell_text = "This cell comment isn't visible (the default)."

    comment = 'Hello.'

    worksheet2.write('C6', cell_text, text_wrap)
    worksheet2.write_comment('C6', comment)

    ###############################################################################
    #
    # Example 3. Demonstrates visible and hidden comments set at the worksheet
    #            level.
    #

    # Set up some formatting.
    worksheet3.set_column('C:C', 25)
    worksheet3.set_row(2, 50)
    worksheet3.set_row(5, 50)
    worksheet3.set_row(8, 50)

    # Make all comments on the worksheet visible.
    worksheet3.show_comments

    cell_text = 'This cell comment is visible, explicitly.'

    comment = 'Hello.'

    worksheet3.write('C3', cell_text, text_wrap)
    worksheet3.write_comment('C3', comment, :visible => 1)

    cell_text =
      'This cell comment is also visible because we used show_comments().'

    comment = 'Hello.'

    worksheet3.write('C6', cell_text, text_wrap)
    worksheet3.write_comment('C6', comment)

    cell_text = 'However, we can still override it locally.'

    comment = 'Hello.'

    worksheet3.write('C9', cell_text, text_wrap)
    worksheet3.write_comment('C9', comment, :visible => 0)

    ###############################################################################
    #
    # Example 4. Demonstrates changes to the comment box dimensions.
    #

    # Set up some formatting.
    worksheet4.set_column('C:C', 25)
    worksheet4.set_row(2,  50)
    worksheet4.set_row(5,  50)
    worksheet4.set_row(8,  50)
    worksheet4.set_row(15, 50)

    worksheet4.show_comments

    cell_text = 'This cell comment is default size.'

    comment = 'Hello.'

    worksheet4.write('C3', cell_text, text_wrap)
    worksheet4.write_comment('C3', comment)

    cell_text = 'This cell comment is twice as wide.'

    comment = 'Hello.'

    worksheet4.write('C6', cell_text, text_wrap)
    worksheet4.write_comment('C6', comment, :x_scale => 2)

    cell_text = 'This cell comment is twice as high.'

    comment = 'Hello.'

    worksheet4.write('C9', cell_text, text_wrap)
    worksheet4.write_comment('C9', comment, :y_scale => 2)

    cell_text = 'This cell comment is scaled in both directions.'

    comment = 'Hello.'

    worksheet4.write('C16', cell_text, text_wrap)
    worksheet4.write_comment('C16', comment, :x_scale => 1.2, :y_scale => 0.8)

    cell_text = 'This cell comment has width and height specified in pixels.'

    comment = 'Hello.'

    worksheet4.write('C19', cell_text, text_wrap)
    worksheet4.write_comment('C19', comment, :width => 200, :height => 20)

    ###############################################################################
    #
    # Example 5. Demonstrates changes to the cell comment position.
    #

    worksheet5.set_column('C:C', 25)
    worksheet5.set_row(2,  50)
    worksheet5.set_row(5,  50)
    worksheet5.set_row(8,  50)
    worksheet5.set_row(11, 50)

    worksheet5.show_comments

    cell_text = 'This cell comment is in the default position.'

    comment = 'Hello.'

    worksheet5.write('C3', cell_text, text_wrap)
    worksheet5.write_comment('C3', comment)

    cell_text = 'This cell comment has been moved to another cell.'

    comment = 'Hello.'

    worksheet5.write('C6', cell_text, text_wrap)
    worksheet5.write_comment('C6', comment, :start_cell => 'E4')

    cell_text = 'This cell comment has been moved to another cell.'

    comment = 'Hello.'

    worksheet5.write('C9', cell_text, text_wrap)
    worksheet5.write_comment('C9', comment, :start_row => 8, :start_col => 4)

    cell_text = 'This cell comment has been shifted within its default cell.'

    comment = 'Hello.'

    worksheet5.write('C12', cell_text, text_wrap)
    worksheet5.write_comment('C12', comment, :x_offset => 30, :y_offset => 12)

    ###############################################################################
    #
    # Example 6. Demonstrates changes to the comment background colour.
    #

    worksheet6.set_column('C:C', 25)
    worksheet6.set_row(2, 50)
    worksheet6.set_row(5, 50)
    worksheet6.set_row(8, 50)

    worksheet6.show_comments

    cell_text = 'This cell comment has a different colour.'

    comment = 'Hello.'

    worksheet6.write('C3', cell_text, text_wrap)
    worksheet6.write_comment('C3', comment, :color => 'green')

    cell_text = 'This cell comment has the default colour.'

    comment = 'Hello.'

    worksheet6.write('C6', cell_text, text_wrap)
    worksheet6.write_comment('C6', comment)

    cell_text = 'This cell comment has a different colour.'

    comment = 'Hello.'

    worksheet6.write('C9', cell_text, text_wrap)
    worksheet6.write_comment('C9', comment, :color => '#FF6600')

    ###############################################################################
    #
    # Example 7. Demonstrates how to set the cell comment author.
    #

    worksheet7.set_column('C:C', 30)
    worksheet7.set_row(2,  50)
    worksheet7.set_row(5,  50)
    worksheet7.set_row(8,  50)

    author = ''
    cell   = 'C3'

    cell_text = "Move the mouse over this cell and you will see 'Cell commented " +
                "by #{author}' (blank) in the status bar at the bottom"

    comment = 'Hello.'

    worksheet7.write(cell, cell_text, text_wrap)
    worksheet7.write_comment(cell, comment)

    author    = 'Ruby'
    cell      = 'C6'
    cell_text = "Move the mouse over this cell and you will see 'Cell commented " +
                "by #{author}' in the status bar at the bottom"

    comment = 'Hello.'

    worksheet7.write(cell, cell_text, text_wrap)
    worksheet7.write_comment(cell, comment, :author => author)

    author    = '€'
    cell      = 'C9'
    cell_text = "Move the mouse over this cell and you will see 'Cell commented " +
                "by #{author}' in the status bar at the bottom"
    comment = 'Hello.'

    worksheet7.write(cell, cell_text, text_wrap)
    worksheet7.write_comment(cell, comment, :author => author)

    ###############################################################################
    #
    # Example 8. Demonstrates the need to explicitly set the row height.
    #

    # Set up some formatting.
    worksheet8.set_column('C:C', 25)
    worksheet8.set_row(2, 80)

    worksheet8.show_comments

    cell_text =
      'The height of this row has been adjusted explicitly using ' +
      'set_row(). The size of the comment box is adjusted ' +
      'accordingly by WriteXLSX.'

    comment = 'Hello.'

    worksheet8.write('C3', cell_text, text_wrap)
    worksheet8.write_comment('C3', comment)

    cell_text =
      'The height of this row has been adjusted by Excel due to the ' +
      'text wrap property being set. Unfortunately this means that ' +
      'the height of the row is unknown to WriteXLSX at ' +
      "run time and thus the comment box is stretched as well.\n\n" +
      'Use set_row() to specify the row height explicitly to avoid ' +
      'this problem.'

    comment = 'Hello.'

    worksheet8.write('C6', cell_text, text_wrap)
    worksheet8.write_comment('C6', comment)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_conditional_format
    @xlsx = 'conditional_format.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet
    worksheet7 = workbook.add_worksheet
    worksheet8 = workbook.add_worksheet

    # Light red fill with dark red text.
    format1 = workbook.add_format(
      :bg_color => '#FFC7CE',
      :color    => '#9C0006'
    )

    # Green fill with dark green text.
    format2 = workbook.add_format(
      :bg_color => '#C6EFCE',
      :color    => '#006100'
    )

    # Some sample data to run the conditional formatting against.
    data = [
      [34, 72,  38, 30, 75, 48, 75, 66, 84, 86],
      [6,  24,  1,  84, 54, 62, 60, 3,  26, 59],
      [28, 79,  97, 13, 85, 93, 93, 22, 5,  14],
      [27, 71,  40, 17, 18, 79, 90, 93, 29, 47],
      [88, 25,  33, 23, 67, 1,  59, 79, 47, 36],
      [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
      [6,  57,  88, 28, 10, 26, 37, 7,  41, 48],
      [52, 78,  1,  96, 26, 45, 47, 33, 96, 36],
      [60, 54,  81, 66, 81, 90, 80, 93, 12, 55],
      [70, 5,   46, 14, 71, 19, 66, 36, 41, 21]
    ]

    ###############################################################################
    #
    # Example 1.
    #
    caption = 'Cells with values >= 50 are in light red. ' +
              'Values < 50 are in light green.'

    # Write the data.
    worksheet1.write('A1', caption)
    worksheet1.write_col('B3', data)

    # Write a conditional format over a range.
    worksheet1.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'cell',
                                        :criteria => '>=',
                                        :value    => 50,
                                        :format   => format1
                                      })

    # Write another conditional format over the same range.
    worksheet1.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'cell',
                                        :criteria => '<',
                                        :value    => 50,
                                        :format   => format2
                                      })

    ###############################################################################
    #
    # Example 2.
    #
    caption = 'Values between 30 and 70 are in light red. ' +
              'Values outside that range are in light green.'

    worksheet2.write('A1', caption)
    worksheet2.write_col('B3', data)

    worksheet2.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'cell',
                                        :criteria => 'between',
                                        :minimum  => 30,
                                        :maximum  => 70,
                                        :format   => format1
                                      })

    worksheet2.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'cell',
                                        :criteria => 'not between',
                                        :minimum  => 30,
                                        :maximum  => 70,
                                        :format   => format2
                                      })

    ###############################################################################
    #
    # Example 3.
    #
    caption = 'Duplicate values are in light red. ' +
              'Unique values are in light green.'

    worksheet3.write('A1', caption)
    worksheet3.write_col('B3', data)

    worksheet3.conditional_formatting('B3:K12',
                                      {
                                        :type   => 'duplicate',
                                        :format => format1
                                      })

    worksheet3.conditional_formatting('B3:K12',
                                      {
                                        :type   => 'unique',
                                        :format => format2
                                      })

    ###############################################################################
    #
    # Example 4.
    #
    caption = 'Above average values are in light red. ' +
              'Below average values are in light green.'

    worksheet4.write('A1', caption)
    worksheet4.write_col('B3', data)

    worksheet4.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'average',
                                        :criteria => 'above',
                                        :format   => format1
                                      })

    worksheet4.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'average',
                                        :criteria => 'below',
                                        :format   => format2
                                      })

    ###############################################################################
    #
    # Example 5.
    #
    caption = 'Top 10 values are in light red. ' +
              'Bottom 10 values are in light green.'

    worksheet5.write('A1', caption)
    worksheet5.write_col('B3', data)

    worksheet5.conditional_formatting('B3:K12',
                                      {
                                        :type   => 'top',
                                        :value  => '10',
                                        :format => format1
                                      })

    worksheet5.conditional_formatting('B3:K12',
                                      {
                                        :type   => 'bottom',
                                        :value  => '10',
                                        :format => format2
                                      })

    ###############################################################################
    #
    # Example 6.
    #
    caption = 'Cells with values >= 50 are in light red. ' +
              'Values < 50 are in light green. Non-contiguous ranges.'

    # Write the data.
    worksheet6.write('A1', caption)
    worksheet6.write_col('B3', data)

    # Write a conditional format over a range.
    worksheet6.conditional_formatting('B3:K6,B9:K12',
                                      {
                                        :type     => 'cell',
                                        :criteria => '>=',
                                        :value    => 50,
                                        :format   => format1
                                      })

    # Write another conditional format over the same range.
    worksheet6.conditional_formatting('B3:K6,B9:K12',
                                      {
                                        :type     => 'cell',
                                        :criteria => '<',
                                        :value    => 50,
                                        :format   => format2
                                      })

    ###############################################################################
    #
    # Example 7.
    #
    caption = 'Examples of color scales and data bars. Default colors.'

    data = 1..12

    worksheet7.write('A1', caption)

    worksheet7.write('B2', "2 Color Scale")
    worksheet7.write_col('B3', data)

    worksheet7.write('D2', "3 Color Scale")
    worksheet7.write_col('D3', data)

    worksheet7.write('F2', "Data Bars")
    worksheet7.write_col('F3', data)

    worksheet7.conditional_formatting('B3:B14',
                                      {
                                        :type => '2_color_scale'
                                      })

    worksheet7.conditional_formatting('D3:D14',
                                      {
                                        :type => '3_color_scale'
                                      })

    worksheet7.conditional_formatting('F3:F14',
                                      {
                                        :type => 'data_bar'
                                      })

    ###############################################################################
    #
    # Example 8.
    #
    caption = 'Examples of color scales and data bars. Modified colors.'

    data = 1..12

    worksheet8.write('A1', caption)

    worksheet8.write('B2', "2 Color Scale")
    worksheet8.write_col('B3', data)

    worksheet8.write('D2', "3 Color Scale")
    worksheet8.write_col('D3', data)

    worksheet8.write('F2', "Data Bars")
    worksheet8.write_col('F3', data)

    worksheet8.conditional_formatting('B3:B14',
                                      {
                                        :type      => '2_color_scale',
                                        :min_color => "#FF0000",
                                        :max_color => "#00FF00"
                                      })

    worksheet8.conditional_formatting('D3:D14',
                                      {
                                        :type      => '3_color_scale',
                                        :min_color => "#C5D9F1",
                                        :mid_color => "#8DB4E3",
                                        :max_color => "#538ED5"
                                      })

    worksheet8.conditional_formatting('F3:F14',
                                      {
                                        :type      => 'data_bar',
                                        :bar_color => '#63C384'
                                      })

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_data_validate
    @xlsx = 'data_validate.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Add a format for the header cells.
    header_format = workbook.add_format(
      :border    => 1,
      :bg_color  => 43,
      :bold      => 1,
      :text_wrap => 1,
      :valign    => 'vcenter',
      :indent    => 1
    )

    # Set up layout of the worksheet.
    worksheet.set_column('A:A', 68)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('D:D', 15)
    worksheet.set_row(0, 36)
    worksheet.set_selection('B3')

    # Write the header cells and some data that will be used in the examples.
    row = 0
    heading1 = 'Some examples of data validation in WriteXLSX'
    heading2 = 'Enter values in this column'
    heading3 = 'Sample Data'

    worksheet.write('A1', heading1, header_format)
    worksheet.write('B1', heading2, header_format)
    worksheet.write('D1', heading3, header_format)

    worksheet.write('D3', ['Integers',   1, 10])
    worksheet.write('D4', ['List data', 'open', 'high', 'close'])
    worksheet.write('D5', ['Formula',   '=AND(F5=50,G5=60)', 50, 60])

    #
    # Example 1. Limiting input to an integer in a fixed range.
    #
    txt = 'Enter an integer between 1 and 10'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'integer',
                                :criteria => 'between',
                                :minimum  => 1,
                                :maximum  => 10
                              })

    #
    # Example 2. Limiting input to an integer outside a fixed range.
    #
    txt = 'Enter an integer that is not between 1 and 10 (using cell references)'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'integer',
                                :criteria => 'not between',
                                :minimum  => '=E3',
                                :maximum  => '=F3'
                              })

    #
    # Example 3. Limiting input to an integer greater than a fixed value.
    #
    txt = 'Enter an integer greater than 0'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'integer',
                                :criteria => '>',
                                :value    => 0
                              })

    #
    # Example 4. Limiting input to an integer less than a fixed value.
    #
    txt = 'Enter an integer less than 10'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'integer',
                                :criteria => '<',
                                :value    => 10
                              })

    #
    # Example 5. Limiting input to a decimal in a fixed range.
    #
    txt = 'Enter a decimal between 0.1 and 0.5'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'decimal',
                                :criteria => 'between',
                                :minimum  => 0.1,
                                :maximum  => 0.5
                              })

    #
    # Example 6. Limiting input to a value in a dropdown list.
    #
    txt = 'Select a value from a drop down list'
    row += 2
    bp = 1
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'list',
                                :source   => %w[open high close]
                              })

    #
    # Example 6. Limiting input to a value in a dropdown list.
    #
    txt = 'Select a value from a drop down list (using a cell range)'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'list',
                                :source   => '=$E$4:$G$4'
                              })

    #
    # Example 7. Limiting input to a date in a fixed range.
    #
    txt = 'Enter a date between 1/1/2008 and 12/12/2008'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'date',
                                :criteria => 'between',
                                :minimum  => '2008-01-01T',
                                :maximum  => '2008-12-12T'
                              })

    #
    # Example 8. Limiting input to a time in a fixed range.
    #
    txt = 'Enter a time between 6:00 and 12:00'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'time',
                                :criteria => 'between',
                                :minimum  => 'T06:00',
                                :maximum  => 'T12:00'
                              })

    #
    # Example 9. Limiting input to a string greater than a fixed length.
    #
    txt = 'Enter a string longer than 3 characters'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'length',
                                :criteria => '>',
                                :value    => 3
                              })

    #
    # Example 10. Limiting input based on a formula.
    #
    txt = 'Enter a value if the following is true "=AND(F5=50,G5=60)"'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate => 'custom',
                                :value    => '=AND(F5=50,G5=60)'
                              })

    #
    # Example 11. Displaying and modify data validation messages.
    #
    txt = 'Displays a message when you select the cell'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate      => 'integer',
                                :criteria      => 'between',
                                :minimum       => 1,
                                :maximum       => 100,
                                :input_title   => 'Enter an integer:',
                                :input_message => 'between 1 and 100'
                              })

    #
    # Example 12. Displaying and modify data validation messages.
    #
    txt = 'Display a custom error message when integer isn\'t between 1 and 100'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate      => 'integer',
                                :criteria      => 'between',
                                :minimum       => 1,
                                :maximum       => 100,
                                :input_title   => 'Enter an integer:',
                                :input_message => 'between 1 and 100',
                                :error_title   => 'Input value is not valid!',
                                :error_message => 'It should be an integer between 1 and 100'
                              })

    #
    # Example 13. Displaying and modify data validation messages.
    #
    txt = 'Display a custom information message when integer isn\'t between 1 and 100'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate      => 'integer',
                                :criteria      => 'between',
                                :minimum       => 1,
                                :maximum       => 100,
                                :input_title   => 'Enter an integer:',
                                :input_message => 'between 1 and 100',
                                :error_title   => 'Input value is not valid!',
                                :error_message => 'It should be an integer between 1 and 100',
                                :error_type    => 'information'
                              })

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_date_time
    @xlsx = 'date_time.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Expand the first column so that the date is visible.
    worksheet.set_column('A:B', 30)

    # Write the column headers.
    worksheet.write('A1', 'Formatted date', bold)
    worksheet.write('B1', 'Format',         bold)

    # Examples date and time formats. In the outpu file compare how changing
    # the format codes change the appearance of the date.
    #
    date_formats = [
      'dd/mm/yy',
      'mm/dd/yy',
      '',
      'd mm yy',
      'dd mm yy',
      '',
      'dd m yy',
      'dd mm yy',
      'dd mmm yy',
      'dd mmmm yy',
      '',
      'dd mm y',
      'dd mm yyy',
      'dd mm yyyy',
      '',
      'd mmmm yyyy',
      '',
      'dd/mm/yy',
      'dd/mm/yy hh:mm',
      'dd/mm/yy hh:mm:ss',
      'dd/mm/yy hh:mm:ss.000',
      '',
      'hh:mm',
      'hh:mm:ss',
      'hh:mm:ss.000'
    ]

    # Write the same date and time using each of the above formats. The empty
    # string formats create a blank line to make the example clearer.
    #
    row = 0
    date_formats.each do |date_format|
      row += 1
      next if date_format == ''

      # Create a format for the date or time.
      format = workbook.add_format(
        :num_format => date_format,
        :align      => 'left'
      )

      # Write the same date using different formats.
      worksheet.write_date_time(row, 0, '2004-08-01T12:30:45.123', format)
      worksheet.write(row, 1, date_format)
    end

    # The following is an example of an invalid date. It is writen as a string
    # instead of a number. This is also Excel's default behaviour.
    #
    row += 2
    worksheet.write_date_time(row, 0, '2004-13-01T12:30:45.123')
    worksheet.write(row, 1, 'Invalid date. Written as string.', bold)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_defined_name
    @xlsx = 'defined_name.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    # Define some global/workbook names.
    workbook.define_name('Exchange_rate', '=0.96')
    workbook.define_name('Sales',         '=Sheet1!$G$1:$H$10')

    # Define a local/worksheet name.
    workbook.define_name('Sheet2!Sales', '=Sheet2!$G$1:$G$10')

    # Write some text in the file and one of the defined names in a formula.
    workbook.worksheets.each do |worksheet|
      worksheet.set_column('A:A', 45)
      worksheet.write('A1', 'This worksheet contains some defined names.')
      worksheet.write('A2', 'See Formulas -> Name Manager above.')
      worksheet.write('A3', 'Example formula in cell B3 ->')

      worksheet.write('B3', '=Exchange_rate')
    end

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_demo
    @xlsx = 'demo.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet  = workbook.add_worksheet('Demo')
    worksheet2 = workbook.add_worksheet('Another sheet')
    worksheet3 = workbook.add_worksheet('And another')

    bold = workbook.add_format(:bold => 1)

    #######################################################################
    #
    # Write a general heading
    #
    worksheet.set_column('A:A', 36, bold)
    worksheet.set_column('B:B', 20)
    worksheet.set_row(0, 40)

    heading = workbook.add_format(
      :bold  => 1,
      :color => 'blue',
      :size  => 16,
      :merge => 1,
      :align => 'vcenter'
    )

    hyperlink_format = workbook.add_format(
      :color     => 'blue',
      :underline => 1
    )

    headings = ['Features of WriteXLSX', '']
    worksheet.write_row('A1', headings, heading)

    #######################################################################
    #
    # Some text examples
    #
    text_format = workbook.add_format(
      :bold   => 1,
      :italic => 1,
      :color  => 'red',
      :size   => 18,
      :font   => 'Lucida Calligraphy'
    )

    worksheet.write('A2', "Text")
    worksheet.write('B2', "Hello Excel")
    worksheet.write('A3', "Formatted text")
    worksheet.write('B3', "Hello Excel", text_format)
    worksheet.write('A4', "Unicode text")
    worksheet.write('B4', "А Б В Г Д")

    #######################################################################
    #
    # Some numeric examples
    #
    num1_format = workbook.add_format(:num_format => '$#,##0.00')
    num2_format = workbook.add_format(:num_format => ' d mmmm yyy')

    worksheet.write('A5', "Numbers")
    worksheet.write('B5', 1234.56)
    worksheet.write('A6', "Formatted numbers")
    worksheet.write('B6', 1234.56, num1_format)
    worksheet.write('A7', "Formatted numbers")
    worksheet.write('B7', 37257, num2_format)

    #######################################################################
    #
    # Formulae
    #
    worksheet.set_selection('B8')
    worksheet.write('A8', 'Formulas and functions, "=SIN(PI()/4)"')
    worksheet.write('B8', '=SIN(PI()/4)')

    #######################################################################
    #
    # Hyperlinks
    #
    worksheet.write('A9', "Hyperlinks")
    worksheet.write('B9', 'http://www.ruby-lang.org/', hyperlink_format)

    #######################################################################
    #
    # Images
    #
    worksheet.write('A10', "Images")
    worksheet.insert_image(
      'B10', File.join(@test_dir, 'republic.png'),
      :x_offset => 16, :y_offset => 8
    )

    #######################################################################
    #
    # Misc
    #
    worksheet.write('A18', "Page/printer setup")
    worksheet.write('A19', "Multiple worksheets")

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_diag_border
    @xlsx = 'diag_border.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    format1 = workbook.add_format(:diag_type => 1)
    format2 = workbook.add_format(:diag_type => 2)
    format3 = workbook.add_format(:diag_type => 3)

    format4 = workbook.add_format(
      :diag_type   => 3,
      :diag_border => 7,
      :diag_color  => 'red'
    )

    worksheet.write('B3',  'Text', format1)
    worksheet.write('B6',  'Text', format2)
    worksheet.write('B9',  'Text', format3)
    worksheet.write('B12', 'Text', format4)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_fit_to_pages
    @xlsx = 'fit_to_pages.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet

    worksheet1.write(0, 0, "fit_to_pages(1, 1)")
    worksheet1.fit_to_pages(1, 1)   # Fit to 1x1 pages

    worksheet2.write(0, 0, "fit_to_pages(2, 1)")
    worksheet2.fit_to_pages(2, 1)   # Fit to 2x1 pages

    worksheet3.write(0, 0, "fit_to_pages(1, 2)")
    worksheet3.fit_to_pages(1, 2)   # Fit to 1x2 pages

    worksheet4.write(0, 0, "fit_to_pages(1, 0)")
    worksheet4.fit_to_pages(1, 0)   # 1 page wide and as long as necessary

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_formats
    @xlsx = 'formats.xlsx'
    workbook = WriteXLSX.new(@io)

    # Some common formats
    center = workbook.add_format(:align => 'center')
    heading = workbook.add_format(:align => 'center', :bold => 1)

    # The named colors
    colors = {
      0x08 => 'black',
      0x0C => 'blue',
      0x10 => 'brown',
      0x0F => 'cyan',
      0x17 => 'gray',
      0x11 => 'green',
      0x0B => 'lime',
      0x0E => 'magenta',
      0x12 => 'navy',
      0x35 => 'orange',
      0x21 => 'pink',
      0x14 => 'purple',
      0x0A => 'red',
      0x16 => 'silver',
      0x09 => 'white',
      0x0D => 'yellow'
    }

    ######################################################################
    #
    # Intro.
    #
    def intro(workbook, _center, _heading, _colors)
      worksheet = workbook.add_worksheet('Introduction')

      worksheet.set_column(0, 0, 60)

      format = workbook.add_format
      format.set_bold
      format.set_size(14)
      format.set_color('blue')
      format.set_align('center')

      format2 = workbook.add_format
      format2.set_bold
      format2.set_color('blue')

      format3 = workbook.add_format(
        :color     => 'blue',
        :underline => 1
      )

      worksheet.write(2, 0, 'This workbook demonstrates some of', format)
      worksheet.write(3, 0, 'the formatting options provided by', format)
      worksheet.write(4, 0, 'the Excel::Writer::XLSX module.',    format)
      worksheet.write('A7', 'Sections:', format2)

      worksheet.write('A8', "internal:Fonts!A1", 'Fonts', format3)

      worksheet.write('A9', "internal:'Named colors'!A1",
                      'Named colors', format3)

      worksheet.write(
        'A10',
        "internal:'Standard colors'!A1",
        'Standard colors', format3
      )

      worksheet.write(
        'A11',
        "internal:'Numeric formats'!A1",
        'Numeric formats', format3
      )

      worksheet.write('A12', "internal:Borders!A1", 'Borders', format3)
      worksheet.write('A13', "internal:Patterns!A1", 'Patterns', format3)
      worksheet.write('A14', "internal:Alignment!A1", 'Alignment', format3)
      worksheet.write('A15', "internal:Miscellaneous!A1", 'Miscellaneous',
                      format3)
    end

    ######################################################################
    #
    # Demonstrate the named colors.
    #
    def named_colors(workbook, center, heading, colors)
      worksheet = workbook.add_worksheet('Named colors')

      worksheet.set_column(0, 3, 15)

      worksheet.write(0, 0, "Index", heading)
      worksheet.write(0, 1, "Index", heading)
      worksheet.write(0, 2, "Name",  heading)
      worksheet.write(0, 3, "Color", heading)

      i = 1

      [33, 11, 53, 17, 22, 18, 13, 16, 23, 9, 12, 15, 14, 20, 8, 10].each do |index|
        color = colors[index]
        format = workbook.add_format(
          :bg_color => color,
          :pattern  => 1,
          :border   => 1
        )

        worksheet.write(i + 1, 0, index, center)
        worksheet.write(i + 1, 1, sprintf("0x%02X", index), center)
        worksheet.write(i + 1, 2, color, center)
        worksheet.write(i + 1, 3, '',     format)
        i += 1
      end
    end

    ######################################################################
    #
    # Demonstrate the standard Excel colors in the range 8..63.
    #
    def standard_colors(workbook, center, heading, colors)
      worksheet = workbook.add_worksheet('Standard colors')

      worksheet.set_column(0, 3, 15)

      worksheet.write(0, 0, "Index", heading)
      worksheet.write(0, 1, "Index", heading)
      worksheet.write(0, 2, "Color", heading)
      worksheet.write(0, 3, "Name",  heading)

      (8..63).each do |i|
        format = workbook.add_format(
          :bg_color => i,
          :pattern  => 1,
          :border   => 1
        )

        worksheet.write((i - 7), 0, i, center)
        worksheet.write((i - 7), 1, sprintf("0x%02X", i), center)
        worksheet.write((i - 7), 2, '', format)

        # Add the  color names
        worksheet.write((i - 7), 3, colors[i], center) if colors[i]
      end
    end

    ######################################################################
    #
    # Demonstrate the standard numeric formats.
    #
    def numeric_formats(workbook, center, heading, _colors)
      worksheet = workbook.add_worksheet('Numeric formats')

      worksheet.set_column(0, 4, 15)
      worksheet.set_column(5, 5, 45)

      worksheet.write(0, 0, "Index",       heading)
      worksheet.write(0, 1, "Index",       heading)
      worksheet.write(0, 2, "Unformatted", heading)
      worksheet.write(0, 3, "Formatted",   heading)
      worksheet.write(0, 4, "Negative",    heading)
      worksheet.write(0, 5, "Format",      heading)

      formats = []
      formats << [0x00, 1234.567,   0,         'General']
      formats << [0x01, 1234.567,   0,         '0']
      formats << [0x02, 1234.567,   0,         '0.00']
      formats << [0x03, 1234.567,   0,         '#,##0']
      formats << [0x04, 1234.567,   0,         '#,##0.00']
      formats << [0x05, 1234.567,   -1234.567, '($#,##0_);($#,##0)']
      formats << [0x06, 1234.567,   -1234.567, '($#,##0_);[Red]($#,##0)']
      formats << [0x07, 1234.567,   -1234.567, '($#,##0.00_);($#,##0.00)']
      formats << [0x08, 1234.567,   -1234.567, '($#,##0.00_);[Red]($#,##0.00)']
      formats << [0x09, 0.567,      0,         '0%']
      formats << [0x0a, 0.567,      0,         '0.00%']
      formats << [0x0b, 1234.567,   0,         '0.00E+00']
      formats << [0x0c, 0.75,       0,         '# ?/?']
      formats << [0x0d, 0.3125,     0,         '# ??/??']
      formats << [0x0e, 36892.521,  0,         'm/d/yy']
      formats << [0x0f, 36892.521,  0,         'd-mmm-yy']
      formats << [0x10, 36892.521,  0,         'd-mmm']
      formats << [0x11, 36892.521,  0,         'mmm-yy']
      formats << [0x12, 36892.521,  0,         'h:mm AM/PM']
      formats << [0x13, 36892.521,  0,         'h:mm:ss AM/PM']
      formats << [0x14, 36892.521,  0,         'h:mm']
      formats << [0x15, 36892.521,  0,         'h:mm:ss']
      formats << [0x16, 36892.521,  0,         'm/d/yy h:mm']
      formats << [0x25, 1234.567,   -1234.567, '(#,##0_);(#,##0)']
      formats << [0x26, 1234.567,   -1234.567, '(#,##0_);[Red](#,##0)']
      formats << [0x27, 1234.567,   -1234.567, '(#,##0.00_);(#,##0.00)']
      formats << [0x28, 1234.567,   -1234.567, '(#,##0.00_);[Red](#,##0.00)']
      formats << [0x29, 1234.567,   -1234.567, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)']
      formats << [0x2a, 1234.567,   -1234.567, '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)']
      formats << [0x2b, 1234.567,   -1234.567, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)']
      formats << [0x2c, 1234.567,   -1234.567, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)']
      formats << [0x2d, 36892.521,  0,         'mm:ss']
      formats << [0x2e, 3.0153,     0,         '[h]:mm:ss']
      formats << [0x2f, 36892.521,  0,         'mm:ss.0']
      formats << [0x30, 1234.567,   0,         '##0.0E+0']
      formats << [0x31, 1234.567,   0,         '@']

      i = 0
      formats.each do |format|
        style = workbook.add_format
        style.set_num_format(format[0])

        i += 1
        worksheet.write(i, 0, format[0], center)
        worksheet.write(i, 1, sprintf("0x%02X", format[0]), center)
        worksheet.write(i, 2, format[1], center)
        worksheet.write(i, 3, format[1], style)

        worksheet.write(i, 4, format[2], style) if format[2] != 0

        worksheet.write_string(i, 5, format[3])
      end
    end

    ######################################################################
    #
    # Demonstrate the font options.
    #
    def fonts(workbook, _center, heading, _colors)
      worksheet = workbook.add_worksheet('Fonts')

      worksheet.set_column(0, 0, 30)
      worksheet.set_column(1, 1, 10)

      worksheet.write(0, 0, "Font name", heading)
      worksheet.write(0, 1, "Font size", heading)

      fonts = []
      fonts << [10, 'Arial']
      fonts << [12, 'Arial']
      fonts << [14, 'Arial']
      fonts << [12, 'Arial Black']
      fonts << [12, 'Arial Narrow']
      fonts << [12, 'Century Schoolbook']
      fonts << [12, 'Courier']
      fonts << [12, 'Courier New']
      fonts << [12, 'Garamond']
      fonts << [12, 'Impact']
      fonts << [12, 'Lucida Handwriting']
      fonts << [12, 'Times New Roman']
      fonts << [12, 'Symbol']
      fonts << [12, 'Wingdings']
      fonts << [12, 'A font that doesn\'t exist']

      i = 0
      fonts.each do |font|
        format = workbook.add_format

        format.set_size(font[0])
        format.set_font(font[1])

        i += 1
        worksheet.write(i, 0, font[1], format)
        worksheet.write(i, 1, font[0], format)
      end
    end

    ######################################################################
    #
    # Demonstrate the standard Excel border styles.
    #
    def borders(workbook, center, heading, _colors)
      worksheet = workbook.add_worksheet('Borders')

      worksheet.set_column(0, 4, 10)
      worksheet.set_column(5, 5, 40)

      worksheet.write(0, 0, "Index",                                heading)
      worksheet.write(0, 1, "Index",                                heading)
      worksheet.write(0, 3, "Style",                                heading)
      worksheet.write(0, 5, "The style is highlighted in red for ", heading)
      worksheet.write(1, 5, "emphasis, the default color is black.",
                      heading)

      14.times do |i|
        format = workbook.add_format
        format.set_border(i)
        format.set_border_color('red')
        format.set_align('center')

        worksheet.write((2 * (i + 1)), 0, i, center)
        worksheet.write((2 * (i + 1)),
                        1, sprintf("0x%02X", i), center)

        worksheet.write((2 * (i + 1)), 3, "Border", format)
      end

      worksheet.write(30, 0, "Diag type",             heading)
      worksheet.write(30, 1, "Index",                 heading)
      worksheet.write(30, 3, "Style",                 heading)
      worksheet.write(30, 5, "Diagonal Boder styles", heading)

      (1..3).each do |i|
        format = workbook.add_format
        format.set_diag_type(i)
        format.set_diag_border(1)
        format.set_diag_color('red')
        format.set_align('center')

        worksheet.write((2 * (i + 15)), 0, i, center)
        worksheet.write((2 * (i + 15)),
                        1, sprintf("0x%02X", i), center)

        worksheet.write((2 * (i + 15)), 3, "Border", format)
      end
    end

    ######################################################################
    #
    # Demonstrate the standard Excel cell patterns.
    #
    def patterns(workbook, center, heading, _colors)
      worksheet = workbook.add_worksheet('Patterns')

      worksheet.set_column(0, 4, 10)
      worksheet.set_column(5, 5, 50)

      worksheet.write(0, 0, "Index",   heading)
      worksheet.write(0, 1, "Index",   heading)
      worksheet.write(0, 3, "Pattern", heading)

      worksheet.write(0, 5, "The background colour has been set to silver.",
                      heading)
      worksheet.write(1, 5, "The foreground colour has been set to green.",
                      heading)

      19.times do |i|
        format = workbook.add_format

        format.set_pattern(i)
        format.set_bg_color('silver')
        format.set_fg_color('green')
        format.set_align('center')

        worksheet.write((2 * (i + 1)), 0, i, center)
        worksheet.write((2 * (i + 1)),
                        1, sprintf("0x%02X", i), center)

        worksheet.write((2 * (i + 1)), 3, "Pattern", format)

        if i == 1
          worksheet.write((2 * (i + 1)),
                          5, "This is solid colour, the most useful pattern.", heading)
        end
      end
    end

    ######################################################################
    #
    # Demonstrate the standard Excel cell alignments.
    #
    def alignment(workbook, _center, heading, _colors)
      worksheet = workbook.add_worksheet('Alignment')

      worksheet.set_column(0, 7, 12)
      worksheet.set_row(0, 40)
      worksheet.set_selection(7, 0)

      format01 = workbook.add_format
      format02 = workbook.add_format
      format03 = workbook.add_format
      format04 = workbook.add_format
      format05 = workbook.add_format
      format06 = workbook.add_format
      format07 = workbook.add_format
      format08 = workbook.add_format
      format09 = workbook.add_format
      format10 = workbook.add_format
      format11 = workbook.add_format
      format12 = workbook.add_format
      format13 = workbook.add_format
      format14 = workbook.add_format
      format15 = workbook.add_format
      format16 = workbook.add_format
      format17 = workbook.add_format

      format02.set_align('top')
      format03.set_align('bottom')
      format04.set_align('vcenter')
      format05.set_align('vjustify')
      format06.set_text_wrap

      format07.set_align('left')
      format08.set_align('right')
      format09.set_align('center')
      format10.set_align('fill')
      format11.set_align('justify')
      format12.set_merge

      format13.set_rotation(45)
      format14.set_rotation(-45)
      format15.set_rotation(270)

      format16.set_shrink
      format17.set_indent(1)

      worksheet.write(0, 0, 'Vertical',   heading)
      worksheet.write(0, 1, 'top',        format02)
      worksheet.write(0, 2, 'bottom',     format03)
      worksheet.write(0, 3, 'vcenter',    format04)
      worksheet.write(0, 4, 'vjustify',   format05)
      worksheet.write(0, 5, "text\nwrap", format06)

      worksheet.write(2, 0, 'Horizontal', heading)
      worksheet.write(2, 1, 'left',       format07)
      worksheet.write(2, 2, 'right',      format08)
      worksheet.write(2, 3, 'center',     format09)
      worksheet.write(2, 4, 'fill',       format10)
      worksheet.write(2, 5, 'justify',    format11)

      worksheet.write(3, 1, 'merge', format12)
      worksheet.write(3, 2, '',      format12)

      worksheet.write(3, 3, 'Shrink ' * 3, format16)
      worksheet.write(3, 4, 'Indent',      format17)

      worksheet.write(5, 0, 'Rotation',   heading)
      worksheet.write(5, 1, 'Rotate 45',  format13)
      worksheet.write(6, 1, 'Rotate -45', format14)
      worksheet.write(7, 1, 'Rotate 270', format15)
    end

    ######################################################################
    #
    # Demonstrate other miscellaneous features.
    #
    def misc(workbook, _center, _heading, _colors)
      worksheet = workbook.add_worksheet('Miscellaneous')

      worksheet.set_column(2, 2, 25)

      format01 = workbook.add_format
      format02 = workbook.add_format
      format03 = workbook.add_format
      format04 = workbook.add_format
      format05 = workbook.add_format
      format06 = workbook.add_format
      format07 = workbook.add_format

      format01.set_underline(0x01)
      format02.set_underline(0x02)
      format03.set_underline(0x21)
      format04.set_underline(0x22)
      format05.set_font_strikeout
      format06.set_font_outline
      format07.set_font_shadow

      worksheet.write(1,  2, 'Underline  0x01',          format01)
      worksheet.write(3,  2, 'Underline  0x02',          format02)
      worksheet.write(5,  2, 'Underline  0x21',          format03)
      worksheet.write(7,  2, 'Underline  0x22',          format04)
      worksheet.write(9,  2, 'Strikeout',                format05)
      worksheet.write(11, 2, 'Outline (Macintosh only)', format06)
      worksheet.write(13, 2, 'Shadow (Macintosh only)',  format07)
    end

    # Call these subroutines to demonstrate different formatting options
    intro(workbook, center, heading, colors)
    fonts(workbook, center, heading, colors)
    named_colors(workbook, center, heading, colors)
    standard_colors(workbook, center, heading, colors)
    numeric_formats(workbook, center, heading, colors)
    borders(workbook, center, heading, colors)
    patterns(workbook, center, heading, colors)
    alignment(workbook, center, heading, colors)
    misc(workbook, center, heading, colors)
    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_headers
    @xlsx = 'headers.xlsx'
    workbook = WriteXLSX.new(@io)
    preview  = 'Select Print Preview to see the header and footer'

    ######################################################################
    #
    # A simple example to start
    #
    worksheet1 = workbook.add_worksheet('Simple')
    header1    = '&CHere is some centred text.'
    footer1    = '&LHere is some left aligned text.'

    worksheet1.set_header(header1)
    worksheet1.set_footer(footer1)

    worksheet1.set_column('A:A', 50)
    worksheet1.write('A1', preview)

    ######################################################################
    #
    # This is an example of some of the header/footer variables.
    #
    worksheet2 = workbook.add_worksheet('Variables')
    header2    = '&LPage &P of &N' + '&CFilename: &F' + '&RSheetname: &A'
    footer2    = '&LCurrent date: &D' + '&RCurrent time: &T'

    worksheet2.set_header(header2)
    worksheet2.set_footer(footer2)

    worksheet2.set_column('A:A', 50)
    worksheet2.write('A1',  preview)
    worksheet2.write('A21', 'Next sheet')
    worksheet2.set_h_pagebreaks(20)

    ######################################################################
    #
    # This example shows how to use more than one font
    #
    worksheet3 = workbook.add_worksheet('Mixed fonts')
    header3    = '&C&"Courier New,Bold"Hello &"Arial,Italic"World'
    footer3    = '&C&"Symbol"e&"Arial" = mc&X2'

    worksheet3.set_header(header3)
    worksheet3.set_footer(footer3)

    worksheet3.set_column('A:A', 50)
    worksheet3.write('A1', preview)

    ######################################################################
    #
    # Example of line wrapping
    #
    worksheet4 = workbook.add_worksheet('Word wrap')
    header4    = "&CHeading 1\nHeading 2"

    worksheet4.set_header(header4)

    worksheet4.set_column('A:A', 50)
    worksheet4.write('A1', preview)

    ######################################################################
    #
    # Example of inserting a literal ampersand &
    #
    worksheet5 = workbook.add_worksheet('Ampersand')
    header5    = '&CCuriouser && Curiouser - Attorneys at Law'

    worksheet5.set_header(header5)

    worksheet5.set_column('A:A', 50)
    worksheet5.write('A1', preview)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_hide_first_sheet
    @xlsx = 'hide_first_sheet.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet2.activate
    worksheet1.hide

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_hide_sheet
    @xlsx = 'hide_sheet.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.set_column('A:A', 30)
    worksheet2.set_column('A:A', 30)
    worksheet3.set_column('A:A', 30)

    # Sheet2 won't be visible until it is unhidden in Excel.
    worksheet2.hide

    worksheet1.write(0, 0, 'Sheet2 is hidden')
    worksheet2.write(0, 0, "Now it's my turn to find you.")
    worksheet3.write(0, 0, 'Sheet2 is hidden')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_hyperlink
    @xlsx = 'hyperlink.xlsx'
    # Create a new workbook and add a worksheet
    workbook = WriteXLSX.new(@io)

    worksheet = workbook.add_worksheet('Hyperlinks')

    # Format the first column
    worksheet.set_column('A:A', 30)
    worksheet.set_selection('B1')

    # Add the standard url link format.
    url_format = workbook.add_format(
      :color     => 'blue',
      :underline => 1
    )

    # Add a sample format.
    red_format = workbook.add_format(
      :color     => 'red',
      :bold      => 1,
      :underline => 1,
      :size      => 12
    )

    # Add an alternate description string to the URL.
    str = 'Perl home.'

    # Add a "tool tip" to the URL.
    tip = 'Get the latest Perl news here.'

    # Write some hyperlinks
    worksheet.write('A1', 'http://www.perl.com/', url_format)
    worksheet.write('A3', 'http://www.perl.com/', url_format, str)
    worksheet.write('A5', 'http://www.perl.com/', url_format, str, tip)
    worksheet.write('A7', 'http://www.perl.com/', red_format)
    worksheet.write('A9', 'mailto:jmcnamara@cpan.org', url_format, 'Mail me')

    # Write a URL that isn't a hyperlink
    worksheet.write_string('A11', 'http://www.perl.com/')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_indent
    @xlsx = 'indent.xlsx'
    workbook = WriteXLSX.new(@io)

    worksheet = workbook.add_worksheet
    indent1   = workbook.add_format(:indent => 1)
    indent2   = workbook.add_format(:indent => 2)

    worksheet.set_column('A:A', 40)

    worksheet.write('A1', "This text is indented 1 level",  indent1)
    worksheet.write('A2', "This text is indented 2 levels", indent2)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_ignore_errors
    @xlsx = 'ignore_errors.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Write strings that looks like numbers. This will cause an Excel warning.
    worksheet.write_string('C2', '123')
    worksheet.write_string('C3', '123')

    # Write a divide by zero formula. This will also cause an Excel warning.
    worksheet.write_formula('C5', '=1/0')
    worksheet.write_formula('C6', '=1/0')

    # Turn off some of the warnings:
    worksheet.ignore_errors(
      :number_stored_as_text => 'C3',
      :eval_error            => 'C6'
    )

    # Write some descriptions for the cells and make the column wider for clarity.
    worksheet.set_column('B:B', 16)
    worksheet.write('B2', 'Warning:')
    worksheet.write('B3', 'Warning turned off:')
    worksheet.write('B5', 'Warning:')
    worksheet.write('B6', 'Warning turned off:')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_keep_leading_zoros
    @xlsx = 'keep_leading_zeros.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.keep_leading_zeros(true)
    worksheet.write('A1', '001')
    worksheet.write('B1', 'written as string.')
    worksheet.write('A2', '012')
    worksheet.write('B2', 'written as string.')
    worksheet.write('A3', '123')
    worksheet.write('B3', 'written as number.')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_merge1
    @xlsx = 'merge1.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    worksheet.set_column('B:D', 20)
    worksheet.set_row(2, 30)

    # Create a merge format
    format = workbook.add_format(:center_across => 1)

    # Only one cell should contain text, the others should be blank.
    worksheet.write(2, 1, "Center across selection", format)
    worksheet.write_blank(2, 2, format)
    worksheet.write_blank(2, 3, format)
    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_merge2
    # Create a new workbook and add a worksheet
    @xlsx = 'merge2.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    worksheet.set_column(1, 2, 30)
    worksheet.set_row(2, 40)

    # Create a merged format
    format = workbook.add_format(
      :center_across => 1,
      :bold          => 1,
      :size          => 15,
      :pattern       => 1,
      :border        => 6,
      :color         => 'white',
      :fg_color      => 'green',
      :border_color  => 'yellow',
      :align         => 'vcenter'
    )

    # Only one cell should contain text, the others should be blank.
    worksheet.write(2, 1, "Center across selection", format)
    worksheet.write_blank(2, 2, format)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_merge3
    @xlsx = 'merge3.xlsx'

    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    [3, 6, 7].each { |row| worksheet.set_row(row, 30) }
    worksheet.set_column('B:D', 20)

    ###############################################################################
    #
    # Example: Merge cells containing a hyperlink using merge_range().
    #
    format = workbook.add_format(
      :border    => 1,
      :underline => 1,
      :color     => 'blue',
      :align     => 'center',
      :valign    => 'vcenter'
    )

    # Merge 3 cells
    worksheet.merge_range('B4:D4', 'http://www.perl.com', format)

    # Merge 3 cells over two rows
    worksheet.merge_range('B7:D8', 'http://www.perl.com', format)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_merge4
    @xlsx = 'merge4.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    (1..11).each { |i| worksheet.set_row(i, 30) }
    worksheet.set_column('B:D', 20)

    ###############################################################################
    #
    # Example 1: Text centered vertically and horizontally
    #
    format1 = workbook.add_format(
      :border => 6,
      :bold   => 1,
      :color  => 'red',
      :valign => 'vcenter',
      :align  => 'center'
    )

    worksheet.merge_range('B2:D3', 'Vertical and horizontal', format1)

    ###############################################################################
    #
    # Example 2: Text aligned to the top and left
    #
    format2 = workbook.add_format(
      :border => 6,
      :bold   => 1,
      :color  => 'red',
      :valign => 'top',
      :align  => 'left'
    )

    worksheet.merge_range('B5:D6', 'Aligned to the top and left', format2)

    ###############################################################################
    #
    # Example 3:  Text aligned to the bottom and right
    #
    format3 = workbook.add_format(
      :border => 6,
      :bold   => 1,
      :color  => 'red',
      :valign => 'bottom',
      :align  => 'right'
    )

    worksheet.merge_range('B8:D9', 'Aligned to the bottom and right', format3)

    ###############################################################################
    #
    # Example 4:  Text justified (i.e. wrapped) in the cell
    #
    format4 = workbook.add_format(
      :border => 6,
      :bold   => 1,
      :color  => 'red',
      :valign => 'top',
      :align  => 'justify'
    )

    worksheet.merge_range('B11:D12', 'Justified: ' + ('so on and ' * 18), format4)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_merge5
    @xlsx = 'merge5.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    (3..8).each { |row|  worksheet.set_row(row, 36) }
    [1, 3, 5].each { |col| worksheet.set_column(col, col, 15) }

    ###############################################################################
    #
    # Rotation 1, letters run from top to bottom
    #
    format1 = workbook.add_format(
      :border   => 6,
      :bold     => 1,
      :color    => 'red',
      :valign   => 'vcentre',
      :align    => 'centre',
      :rotation => 270
    )

    worksheet.merge_range('B4:B9', 'Rotation 270', format1)

    ###############################################################################
    #
    # Rotation 2, 90ｰ anticlockwise
    #
    format2 = workbook.add_format(
      :border   => 6,
      :bold     => 1,
      :color    => 'red',
      :valign   => 'vcentre',
      :align    => 'centre',
      :rotation => 90
    )

    worksheet.merge_range('D4:D9', 'Rotation 90°', format2)

    ###############################################################################
    #
    # Rotation 3, 90ｰ clockwise
    #
    format3 = workbook.add_format(
      :border   => 6,
      :bold     => 1,
      :color    => 'red',
      :valign   => 'vcentre',
      :align    => 'centre',
      :rotation => -90
    )

    worksheet.merge_range('F4:F9', 'Rotation -90°', format3)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_merge6
    @xlsx = 'merge6.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    (2..9).each { |row| worksheet.set_row(row, 36) }
    worksheet.set_column('B:D', 25)

    # Format for the merged cells.
    format = workbook.add_format(
      :border => 6,
      :bold   => 1,
      :color  => 'red',
      :size   => 20,
      :valign => 'vcentre',
      :align  => 'left',
      :indent => 1
    )

    ###############################################################################
    #
    # Write an Ascii string.
    #
    worksheet.merge_range('B3:D4', 'ASCII: A simple string', format)

    ###############################################################################
    #
    # Write a UTF-8 Unicode string.
    #
    smiley = '☺'
    worksheet.merge_range('B6:D7', "UTF-8: A Unicode smiley #{smiley}", format)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_outline
    @xlsx = 'outline.xlsx'
    # Create a new workbook and add some worksheets
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet('Outlined Rows')
    worksheet2 = workbook.add_worksheet('Collapsed Rows')
    worksheet3 = workbook.add_worksheet('Outline Columns')
    worksheet4 = workbook.add_worksheet('Outline levels')

    # Add a general format
    bold = workbook.add_format(:bold => 1)

    ###############################################################################
    #
    # Example 1: Create a worksheet with outlined rows. It also includes SUBTOTAL()
    # functions so that it looks like the type of automatic outlines that are
    # generated when you use the Excel Data.SubTotals menu item.
    #

    # For outlines the important parameters are $hidden and $level. Rows with the
    # same $level are grouped together. The group will be collapsed if $hidden is
    # non-zero. $height and $XF are assigned default values if they are nil.
    #
    # The syntax is: set_row($row, $height, $XF, $hidden, $level, $collapsed)
    #
    worksheet1.set_row(1, nil, nil, 0, 2)
    worksheet1.set_row(2, nil, nil, 0, 2)
    worksheet1.set_row(3, nil, nil, 0, 2)
    worksheet1.set_row(4, nil, nil, 0, 2)
    worksheet1.set_row(5, nil, nil, 0, 1)

    worksheet1.set_row(6,  nil, nil, 0, 2)
    worksheet1.set_row(7,  nil, nil, 0, 2)
    worksheet1.set_row(8,  nil, nil, 0, 2)
    worksheet1.set_row(9,  nil, nil, 0, 2)
    worksheet1.set_row(10, nil, nil, 0, 1)

    # Add a column format for clarity
    worksheet1.set_column('A:A', 20)

    # Add the data, labels and formulas
    worksheet1.write('A1', 'Region', bold)
    worksheet1.write('A2', 'North')
    worksheet1.write('A3', 'North')
    worksheet1.write('A4', 'North')
    worksheet1.write('A5', 'North')
    worksheet1.write('A6', 'North Total', bold)

    worksheet1.write('B1', 'Sales', bold)
    worksheet1.write('B2', 1000)
    worksheet1.write('B3', 1200)
    worksheet1.write('B4', 900)
    worksheet1.write('B5', 1200)
    worksheet1.write('B6', '=SUBTOTAL(9,B2:B5)', bold)

    worksheet1.write('A7',  'South')
    worksheet1.write('A8',  'South')
    worksheet1.write('A9',  'South')
    worksheet1.write('A10', 'South')
    worksheet1.write('A11', 'South Total', bold)

    worksheet1.write('B7',  400)
    worksheet1.write('B8',  600)
    worksheet1.write('B9',  500)
    worksheet1.write('B10', 600)
    worksheet1.write('B11', '=SUBTOTAL(9,B7:B10)', bold)

    worksheet1.write('A12', 'Grand Total',         bold)
    worksheet1.write('B12', '=SUBTOTAL(9,B2:B10)', bold)

    ###############################################################################
    #
    # Example 2: Create a worksheet with outlined rows. This is the same as the
    # previous example except that the rows are collapsed.
    # Note: We need to indicate the row that contains the collapsed symbol '+'
    # with the optional parameter, $collapsed.

    # The group will be collapsed if $hidden is non-zero.
    # The syntax is: set_row($row, $height, $XF, $hidden, $level, $collapsed)
    #
    worksheet2.set_row(1, nil, nil, 1, 2)
    worksheet2.set_row(2, nil, nil, 1, 2)
    worksheet2.set_row(3, nil, nil, 1, 2)
    worksheet2.set_row(4, nil, nil, 1, 2)
    worksheet2.set_row(5, nil, nil, 1, 1)

    worksheet2.set_row(6,  nil, nil, 1, 2)
    worksheet2.set_row(7,  nil, nil, 1, 2)
    worksheet2.set_row(8,  nil, nil, 1, 2)
    worksheet2.set_row(9,  nil, nil, 1, 2)
    worksheet2.set_row(10, nil, nil, 1, 1)
    worksheet2.set_row(11, nil, nil, 0, 0, 1)

    # Add a column format for clarity
    worksheet2.set_column('A:A', 20)

    # Add the data, labels and formulas
    worksheet2.write('A1', 'Region', bold)
    worksheet2.write('A2', 'North')
    worksheet2.write('A3', 'North')
    worksheet2.write('A4', 'North')
    worksheet2.write('A5', 'North')
    worksheet2.write('A6', 'North Total', bold)

    worksheet2.write('B1', 'Sales', bold)
    worksheet2.write('B2', 1000)
    worksheet2.write('B3', 1200)
    worksheet2.write('B4', 900)
    worksheet2.write('B5', 1200)
    worksheet2.write('B6', '=SUBTOTAL(9,B2:B5)', bold)

    worksheet2.write('A7',  'South')
    worksheet2.write('A8',  'South')
    worksheet2.write('A9',  'South')
    worksheet2.write('A10', 'South')
    worksheet2.write('A11', 'South Total', bold)

    worksheet2.write('B7',  400)
    worksheet2.write('B8',  600)
    worksheet2.write('B9',  500)
    worksheet2.write('B10', 600)
    worksheet2.write('B11', '=SUBTOTAL(9,B7:B10)', bold)

    worksheet2.write('A12', 'Grand Total',         bold)
    worksheet2.write('B12', '=SUBTOTAL(9,B2:B10)', bold)

    ###############################################################################
    #
    # Example 3: Create a worksheet with outlined columns.
    #
    data = [
      ['Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', ' Total'],
      ['North', 50,    20,    15,    25,    65,    80,    '=SUM(B2:G2)'],
      ['South', 10,    20,    30,    50,    50,    50,    '=SUM(B3:G3)'],
      ['East',  45,    75,    50,    15,    75,    100,   '=SUM(B4:G4)'],
      ['West',  15,    15,    55,    35,    20,    50,    '=SUM(B5:G5)']
    ]

    # Add bold format to the first row
    worksheet3.set_row(0, nil, bold)

    # Syntax: set_column($col1, $col2, $width, $XF, $hidden, $level, $collapsed)
    worksheet3.set_column('A:A', 10, bold)
    worksheet3.set_column('B:G', 5, nil, 0, 1)
    worksheet3.set_column('H:H', 10)

    # Write the data and a formula
    worksheet3.write_col('A1', data)
    worksheet3.write('H6', '=SUM(H2:H5)', bold)

    ###############################################################################
    #
    # Example 4: Show all possible outline levels.
    #
    levels = [
      "Level 1", "Level 2", "Level 3", "Level 4", "Level 5", "Level 6",
      "Level 7", "Level 6", "Level 5", "Level 4", "Level 3", "Level 2",
      "Level 1"
    ]

    worksheet4.write_col('A1', levels)

    worksheet4.set_row(0,  nil, nil, nil, 1)
    worksheet4.set_row(1,  nil, nil, nil, 2)
    worksheet4.set_row(2,  nil, nil, nil, 3)
    worksheet4.set_row(3,  nil, nil, nil, 4)
    worksheet4.set_row(4,  nil, nil, nil, 5)
    worksheet4.set_row(5,  nil, nil, nil, 6)
    worksheet4.set_row(6,  nil, nil, nil, 7)
    worksheet4.set_row(7,  nil, nil, nil, 6)
    worksheet4.set_row(8,  nil, nil, nil, 5)
    worksheet4.set_row(9,  nil, nil, nil, 4)
    worksheet4.set_row(10, nil, nil, nil, 3)
    worksheet4.set_row(11, nil, nil, nil, 2)
    worksheet4.set_row(12, nil, nil, nil, 1)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_outline_collapsed
    @xlsx = 'outline_collapsed.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet('Outlined Rows')
    worksheet2 = workbook.add_worksheet('Collapsed Rows 1')
    worksheet3 = workbook.add_worksheet('Collapsed Rows 2')
    worksheet4 = workbook.add_worksheet('Collapsed Rows 3')
    worksheet5 = workbook.add_worksheet('Outline Columns')
    worksheet6 = workbook.add_worksheet('Collapsed Columns')

    # Add a general format
    bold = workbook.add_format(:bold => 1)

    ###############################################################################
    #
    # Example 1: Create a worksheet with outlined rows. It also includes SUBTOTAL()
    # functions so that it looks like the type of automatic outlines that are
    # generated when you use the Excel Data->SubTotals menu item.
    #

    # The syntax is: set_row(row, height, XF, hidden, level, collapsed)
    worksheet1.set_row(1, nil, nil, 0, 2)
    worksheet1.set_row(2, nil, nil, 0, 2)
    worksheet1.set_row(3, nil, nil, 0, 2)
    worksheet1.set_row(4, nil, nil, 0, 2)
    worksheet1.set_row(5, nil, nil, 0, 1)

    worksheet1.set_row(6,  nil, nil, 0, 2)
    worksheet1.set_row(7,  nil, nil, 0, 2)
    worksheet1.set_row(8,  nil, nil, 0, 2)
    worksheet1.set_row(9,  nil, nil, 0, 2)
    worksheet1.set_row(10, nil, nil, 0, 1)

    # Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet1, bold)

    ###############################################################################
    #
    # Example 2: Create a worksheet with collapsed outlined rows.
    # This is the same as the example 1  except that the all rows are collapsed.
    # Note: We need to indicate the row that contains the collapsed symbol '+' with
    # the optional parameter, collapsed.

    worksheet2.set_row(1, nil, nil, 1, 2)
    worksheet2.set_row(2, nil, nil, 1, 2)
    worksheet2.set_row(3, nil, nil, 1, 2)
    worksheet2.set_row(4, nil, nil, 1, 2)
    worksheet2.set_row(5, nil, nil, 1, 1)

    worksheet2.set_row(6,  nil, nil, 1, 2)
    worksheet2.set_row(7,  nil, nil, 1, 2)
    worksheet2.set_row(8,  nil, nil, 1, 2)
    worksheet2.set_row(9,  nil, nil, 1, 2)
    worksheet2.set_row(10, nil, nil, 1, 1)

    worksheet2.set_row(11, nil, nil, 0, 0, 1)

    # Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet2, bold)

    ###############################################################################
    #
    # Example 3: Create a worksheet with collapsed outlined rows.
    # Same as the example 1  except that the two sub-totals are collapsed.

    worksheet3.set_row(1, nil, nil, 1, 2)
    worksheet3.set_row(2, nil, nil, 1, 2)
    worksheet3.set_row(3, nil, nil, 1, 2)
    worksheet3.set_row(4, nil, nil, 1, 2)
    worksheet3.set_row(5, nil, nil, 0, 1, 1)

    worksheet3.set_row(6,  nil, nil, 1, 2)
    worksheet3.set_row(7,  nil, nil, 1, 2)
    worksheet3.set_row(8,  nil, nil, 1, 2)
    worksheet3.set_row(9,  nil, nil, 1, 2)
    worksheet3.set_row(10, nil, nil, 0, 1, 1)

    # Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet3, bold)

    ###############################################################################
    #
    # Example 4: Create a worksheet with outlined rows.
    # Same as the example 1  except that the two sub-totals are collapsed.

    worksheet4.set_row(1, nil, nil, 1, 2)
    worksheet4.set_row(2, nil, nil, 1, 2)
    worksheet4.set_row(3, nil, nil, 1, 2)
    worksheet4.set_row(4, nil, nil, 1, 2)
    worksheet4.set_row(5, nil, nil, 1, 1, 1)

    worksheet4.set_row(6,  nil, nil, 1, 2)
    worksheet4.set_row(7,  nil, nil, 1, 2)
    worksheet4.set_row(8,  nil, nil, 1, 2)
    worksheet4.set_row(9,  nil, nil, 1, 2)
    worksheet4.set_row(10, nil, nil, 1, 1, 1)

    worksheet4.set_row(11, nil, nil, 0, 0, 1)

    # Write the sub-total data that is common to the row examples.
    create_sub_totals(worksheet4, bold)

    ###############################################################################
    #
    # Example 5: Create a worksheet with outlined columns.
    #
    data = [
      %w[Month Jan Feb Mar Apr May Jun Total],
      ['North', 50,    20,    15,    25,    65,    80,   '=SUM(B2:G2)'],
      ['South', 10,    20,    30,    50,    50,    50,   '=SUM(B3:G3)'],
      ['East',  45,    75,    50,    15,    75,    100,  '=SUM(B4:G4)'],
      ['West',  15,    15,    55,    35,    20,    50,   '=SUM(B5:G6)']
    ]

    # Add bold format to the first row
    worksheet5.set_row(0, nil, bold)

    # Syntax: set_column(col1, col2, width, XF, hidden, level, collapsed)
    worksheet5.set_column('A:A', 10, bold)
    worksheet5.set_column('B:G', 5, nil, 0, 1)
    worksheet5.set_column('H:H', 10)

    # Write the data and a formula
    worksheet5.write_col('A1', data)
    worksheet5.write('H6', '=SUM(H2:H5)', bold)

    ###############################################################################
    #
    # Example 6: Create a worksheet with collapsed outlined columns.
    # This is the same as the previous example except collapsed columns.

    # Add bold format to the first row
    worksheet6.set_row(0, nil, bold)

    # Syntax: set_column(col1, col2, width, XF, hidden, level, collapsed)
    worksheet6.set_column('A:A', 10, bold)
    worksheet6.set_column('B:G', 5,  nil, 1, 1)
    worksheet6.set_column('H:H', 10, nil, 0, 0, 1)

    # Write the data and a formula
    worksheet6.write_col('A1', data)
    worksheet6.write('H6', '=SUM(H2:H5)', bold)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  #
  # This function will generate the same data and sub-totals on each worksheet.
  #
  def create_sub_totals(worksheet, bold)
    # Add a column format for clarity
    worksheet.set_column('A:A', 20)

    # Add the data, labels and formulas
    worksheet.write('A1', 'Region', bold)
    worksheet.write('A2', 'North')
    worksheet.write('A3', 'North')
    worksheet.write('A4', 'North')
    worksheet.write('A5', 'North')
    worksheet.write('A6', 'North Total', bold)

    worksheet.write('B1', 'Sales', bold)
    worksheet.write('B2', 1000)
    worksheet.write('B3', 1200)
    worksheet.write('B4', 900)
    worksheet.write('B5', 1200)
    worksheet.write('B6', '=SUBTOTAL(9,B2:B5)', bold)

    worksheet.write('A7',  'South')
    worksheet.write('A8',  'South')
    worksheet.write('A9',  'South')
    worksheet.write('A10', 'South')
    worksheet.write('A11', 'South Total', bold)

    worksheet.write('B7',  400)
    worksheet.write('B8',  600)
    worksheet.write('B9',  500)
    worksheet.write('B10', 600)
    worksheet.write('B11', '=SUBTOTAL(9,B7:B10)', bold)

    worksheet.write('A12', 'Grand Total',         bold)
    worksheet.write('B12', '=SUBTOTAL(9,B2:B10)', bold)
  end

  def test_panes
    @xlsx = 'panes.xlsx'
    workbook  = WriteXLSX.new(@io)

    worksheet1 = workbook.add_worksheet('Panes 1')
    worksheet2 = workbook.add_worksheet('Panes 2')
    worksheet3 = workbook.add_worksheet('Panes 3')
    worksheet4 = workbook.add_worksheet('Panes 4')

    # Freeze panes
    worksheet1.freeze_panes(1, 0)    # 1 row

    worksheet2.freeze_panes(0, 1)    # 1 column
    worksheet3.freeze_panes(1, 1)    # 1 row and column

    # Split panes.
    # The divisions must be specified in terms of row and column dimensions.
    # The default row height is 15 and the default column width is 8.43
    #
    worksheet4.split_panes(15, 8.43)    # 1 row and column

    #######################################################################
    #
    # Set up some formatting and text to highlight the panes
    #

    header = workbook.add_format(
      :align    => 'center',
      :valign   => 'vcenter',
      :fg_color => 0x2A
    )

    center = workbook.add_format(:align => 'center')

    #######################################################################
    #
    # Sheet 1
    #

    worksheet1.set_column('A:I', 16)
    worksheet1.set_row(0, 20)
    worksheet1.set_selection('C3')

    9.times { |i| worksheet1.write(0, i, 'Scroll down', header) }
    (1..100).each do |i|
      9.times { |j| worksheet1.write(i, j, i + 1, center) }
    end

    #######################################################################
    #
    # Sheet 2
    #

    worksheet2.set_column('A:A', 16)
    worksheet2.set_selection('C3')

    50.times do |i|
      worksheet2.set_row(i, 15)
      worksheet2.write(i, 0, 'Scroll right', header)
    end

    50.times do |i|
      (1..25).each { |j| worksheet2.write(i, j, j, center) }
    end

    #######################################################################
    #
    # Sheet 3
    #

    worksheet3.set_column('A:Z', 16)
    worksheet3.set_selection('C3')

    worksheet3.write(0, 0, '', header)

    (1..25).each { |i| worksheet3.write(0, i, 'Scroll down', header) }
    (1..49).each { |i| worksheet3.write(i, 0, 'Scroll right', header) }
    (1..49).each do |i|
      (1..25).each { |j| worksheet3.write(i, j, j, center) }
    end

    #######################################################################
    #
    # Sheet 4
    #

    worksheet4.set_selection('C3')

    (1..25).each { |i| worksheet4.write(0, i, 'Scroll', center) }
    (1..49).each { |i| worksheet4.write(i, 0, 'Scroll', center) }
    (1..49).each do |i|
      (1..25).each { |j| worksheet4.write(i, j, j, center) }
    end

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_print_scale
    @xlsx = 'print_scale.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.write(0, 0, "print_scale(100)")
    worksheet1.print_scale = 100

    worksheet2.write(0, 0, "print_scale(50)")
    worksheet2.print_scale = 50

    worksheet3.write(0, 0, "print_scale(200)")
    worksheet3.print_scale = 200

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_properties
    @xlsx = 'properties.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_properties(
      :title    => 'This is an example spreadsheet',
      :subject  => 'With document properties',
      :author   => 'John McNamara',
      :manager  => 'Dr. Heinz Doofenshmirtz',
      :company  => 'of Wolves',
      :category => 'Example spreadsheets',
      :keywords => 'Sample, Example, Properties',
      :comments => 'Created with Perl and Excel::Writer::XLSX',
      :status   => 'Quo'
    )

    worksheet.set_column('A:A', 70)
    worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_protection
    @xlsx = 'protection.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Create some format objects
    unlocked = workbook.add_format(:locked => 0)
    hidden   = workbook.add_format(:hidden => 1)

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
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_regions
    @xlsx = 'regions.xlsx'
    workbook  = WriteXLSX.new(@io)

    # Add some worksheets
    north = workbook.add_worksheet('North')
    south = workbook.add_worksheet('South')
    east  = workbook.add_worksheet('East')
    west  = workbook.add_worksheet('West')

    # Add a Format
    format = workbook.add_format
    format.set_bold
    format.set_color('blue')

    # Add a caption to each worksheet
    workbook.sheets.each do |worksheet|
      worksheet.write(0, 0, 'Sales', format)
    end

    # Write some data
    north.write(0, 1, 200000)
    south.write(0, 1, 100000)
    east.write(0, 1, 150000)
    west.write(0, 1, 100000)

    # Set the active worksheet
    south.activate

    # Set the width of the first column
    south.set_column(0, 0, 20)

    # Set the active cell
    south.set_selection(0, 1)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_rich_strings
    @xlsx = 'rich_strings.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('A:A', 30)

    # Set some formats to use.
    bold   = workbook.add_format(:bold        => 1)
    italic = workbook.add_format(:italic      => 1)
    red    = workbook.add_format(:color       => 'red')
    blue   = workbook.add_format(:color       => 'blue')
    center = workbook.add_format(:align       => 'center')
    superc = workbook.add_format(:font_script => 1)

    # Write some strings with multiple formats.
    worksheet.write_rich_string('A1',
                                'This is ', bold, 'bold', ' and this is ', italic, 'italic')

    worksheet.write_rich_string('A3',
                                'This is ', red, 'red', ' and this is ', blue, 'blue')

    worksheet.write_rich_string('A5',
                                'Some ', bold, 'bold text', ' centered', center)

    worksheet.write_rich_string('A7',
                                italic, 'j = k', superc, '(n-1)', center)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_right_to_left
    @xlsx = 'right_to_left.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet2.right_to_left

    worksheet1.write(0, 0, 'Hello')    #  A1, B1, C1, ...
    worksheet2.write(0, 0, 'Hello')    # ..., C1, B1, A1
    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape1
    @xlsx = 'shape1.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Add a circle, with centered text.
    ellipse = workbook.add_shape(
      :type   => 'ellipse',
      :text   => "Hello\nWorld",
      :width  => 60,
      :height => 60
    )

    worksheet.insert_shape('A1', ellipse, 50, 50)

    # Add a plus sign.
    plus = workbook.add_shape(:type => 'plus', :width => 20, :height => 20)
    worksheet.insert_shape('D8', plus)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape2
    @xlsx = 'shape2.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.hide_gridlines(2)

    plain = workbook.add_shape(
      :type   => 'smileyFace',
      :text   => "Plain",
      :width  => 100,
      :height => 100
    )

    bbformat = workbook.add_format(
      :color => 'red',
      :font  => 'Lucida Calligraphy'
    )

    bbformat.set_bold
    bbformat.set_underline
    bbformat.set_italic

    decor = workbook.add_shape(
      :type        => 'smileyFace',
      :text        => 'Decorated',
      :rotation    => 45,
      :width       => 200,
      :height      => 100,
      :format      => bbformat,
      :line_type   => 'sysDot',
      :line_weight => 3,
      :fill        => 'FFFF00',
      :line        => '3366FF'
    )

    worksheet.insert_shape('A1', plain,  50, 50)
    worksheet.insert_shape('A1', decor, 250, 50)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape3
    @xlsx = 'shape3.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    normal = workbook.add_shape(
      :name   => 'chip',
      :type   => 'diamond',
      :text   => 'Normal',
      :width  => 100,
      :height => 100
    )

    worksheet.insert_shape('A1', normal, 50, 50)
    normal.text = 'Scaled 3w x 2h'
    normal.name = 'Hope'
    worksheet.insert_shape('A1', normal, 250, 50, 3, 2)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape4
    @xlsx = 'shape4.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    worksheet.hide_gridlines(2)

    type = 'rect'
    shape = workbook.add_shape(
      :type   => type,
      :width  => 90,
      :height => 90
    )

    (1..10).each do |n|
      # Change the last 5 rectangles to stars. Previously
      # inserted shapes stay as rectangles.
      type = 'star5' if n == 6
      shape.type = type
      shape.text = "#{type} #{n}"
      worksheet.insert_shape('A1', shape, n * 100, 50)
    end

    stencil = workbook.add_shape(
      :stencil => 1,     # The default.
      :width   => 90,
      :height  => 90,
      :text    => 'started as a box'
    )
    worksheet.insert_shape('A1', stencil, 100, 150)

    stencil.stencil = 0
    worksheet.insert_shape('A1', stencil, 200, 150)
    worksheet.insert_shape('A1', stencil, 300, 150)

    # Ooopa! Changed my mind.
    # Change the rectangle to an ellipse (circle),
    # for the last two shapes.
    stencil.type = 'ellipse'
    stencil.text = 'Now its a circle'

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape5
    @xlsx = 'shape5.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    s1 = workbook.add_shape(
      :type   => 'ellipse',
      :width  => 60,
      :height => 60
    )
    worksheet.insert_shape('A1', s1, 50, 50)

    s2 = workbook.add_shape(
      :type   => 'plus',
      :width  => 20,
      :height => 20
    )
    worksheet.insert_shape('A1', s2, 250, 200)

    # Create a connector to link the two shapes.
    cxn_shape = workbook.add_shape(:type => 'bentConnector3')

    # Link the start of the connector to the right side.
    cxn_shape.start       = s1.id
    cxn_shape.start_index = 4  # 4th connection pt, clockwise from top(0).
    cxn_shape.start_side  = 'b' # r)ight or b)ottom.

    # Link the end of the connector to the left side.
    cxn_shape.end         = s2.id
    cxn_shape.end_index   = 0  # clockwise from top(0).
    cxn_shape.end_side    = 't' # t)op.

    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape6
    @xlsx = 'shape6.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    s1 = workbook.add_shape(
      :type   => 'chevron',
      :width  => 60,
      :height => 60
    )
    worksheet.insert_shape('A1', s1, 50, 50)

    s2 = workbook.add_shape(
      :type   => 'pentagon',
      :width  => 20,
      :height => 20
    )
    worksheet.insert_shape('A1', s2, 250, 200)

    # Create a connector to link the two shapes.
    cxn_shape = workbook.add_shape(:type => 'curvedConnector3')

    # Link the start of the connector to the right side.
    cxn_shape.start       = s1.id
    cxn_shape.start_index = 2  # 2nd connection pt, clockwise from top(0).
    cxn_shape.start_side  = 'r' # r)ight or b)ottom.

    # Link the end of the connector to the left side.
    cxn_shape.end         = s2.id
    cxn_shape.end_index   = 4  # 4th connection pt, clockwise from top(0).
    cxn_shape.end_side    = 'l' # l)eft or t)op.

    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape7
    @xlsx = 'shape7.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Add a circle, with centered text. c is for circle, not center.
    cw = 60
    ch = 60
    cx = 210
    cy = 190

    ellipse = workbook.add_shape(
      :type   => 'ellipse',
      :id     => 2,
      :text   => "Hello\nWorld",
      :width  => cw,
      :height => ch
    )
    worksheet.insert_shape('A1', ellipse, cx, cy)

    # Add a plus sign at 4 different positions around the circle.
    pw = 20
    ph = 20
    px = 120
    py = 250

    plus = workbook.add_shape(
      :type   => 'plus',
      :id     => 3,
      :width  => pw,
      :height => ph
    )

    p1 = worksheet.insert_shape('A1', plus, 350, 350)
    p2 = worksheet.insert_shape('A1', plus, 150, 350)
    p3 = worksheet.insert_shape('A1', plus, 350, 150)
    plus.adjustments = 35  # change shape of plus symbol.
    p4 = worksheet.insert_shape('A1', plus, 150, 150)

    cxn_shape = workbook.add_shape(:type => 'bentConnector3', :fill => 0)

    cxn_shape.start       = ellipse.id
    cxn_shape.start_index = 4   # 4th connection pt, clockwise from top(0).
    cxn_shape.start_side  = 'b' # r)ight or b)ottom.

    cxn_shape.end         = p1.id
    cxn_shape.end_index   = 0
    cxn_shape.end_side    = 't' # l)eft or t)op.
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p2.id
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p3.id
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p4.id
    cxn_shape.adjustments = [-50, 45, 120]
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape8
    @xlsx = 'shape8.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Add a circle, with centered text. c is for circle, not center.
    cw = 60
    ch = 60
    cx = 210
    cy = 190

    ellipse = workbook.add_shape(
      :type   => 'ellipse',
      :id     => 2,
      :text   => "Hello\nWorld",
      :width  => cw,
      :height => ch
    )
    worksheet.insert_shape('A1', ellipse, cx, cy)

    # Add a plus sign at 4 different positionos around the circle.
    pw = 20
    ph = 20
    px = 120
    py = 250

    plus = workbook.add_shape(
      :type   => 'plus',
      :id     => 3,
      :width  => pw,
      :height => ph
    )

    p1 = worksheet.insert_shape('A1', plus, 350, 150)
    p2 = worksheet.insert_shape('A1', plus, 350, 350)
    p3 = worksheet.insert_shape('A1', plus, 150, 350)
    p4 = worksheet.insert_shape('A1', plus, 150, 150)

    cxn_shape = workbook.add_shape(:type => 'bentConnector3', :fill => 0)

    cxn_shape.start       = ellipse.id
    cxn_shape.start_index = 2   # 2nd connection pt, clockwise from top(0).
    cxn_shape.start_side  = 'r' # r)ight or b)ottom.

    cxn_shape.end         = p1.id
    cxn_shape.end_index   = 3   # 3rd connection point on plus, right side
    cxn_shape.end_side    = 'l' # l)eft or t)op.
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p2.id
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p3.id
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    cxn_shape.end = p4.id
    cxn_shape.adjustments = [-50, 45, 120]
    worksheet.insert_shape('A1', cxn_shape, 0, 0)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_shape_all
    @xlsx = 'shape_all.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet  = nil
    last_sheet = ''
    row        = 0

    shapes_list.each_line do |line|
      line = line.chomp
      next unless line =~ /^\w/    # Skip blank lines and comments.

      sheet, name = line.split(/\t/)
      if last_sheet != sheet
        worksheet = workbook.add_worksheet(sheet)
        row       = 2
      end
      last_sheet = sheet
      shape      = workbook.add_shape(
        :type   => name,
        :text   => name,
        :width  => 90,
        :height => 90
      )

      # Connectors can not have labels, so write the connector name in the cell
      # to the left.
      worksheet.write(row, 0, name) if sheet == 'Connector'
      worksheet.insert_shape(row, 2, shape, 0, 0)
      row += 5
    end

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def shapes_list
    <<EOS
Action	actionButtonBackPrevious
Action	actionButtonBeginning
Action	actionButtonBlank
Action	actionButtonDocument
Action	actionButtonEnd
Action	actionButtonForwardNext
Action	actionButtonHelp
Action	actionButtonHome
Action	actionButtonInformation
Action	actionButtonMovie
Action	actionButtonReturn
Action	actionButtonSound
Arrow	bentArrow
Arrow	bentUpArrow
Arrow	circularArrow
Arrow	curvedDownArrow
Arrow	curvedLeftArrow
Arrow	curvedRightArrow
Arrow	curvedUpArrow
Arrow	downArrow
Arrow	leftArrow
Arrow	leftCircularArrow
Arrow	leftRightArrow
Arrow	leftRightCircularArrow
Arrow	leftRightUpArrow
Arrow	leftUpArrow
Arrow	notchedRightArrow
Arrow	quadArrow
Arrow	rightArrow
Arrow	stripedRightArrow
Arrow	swooshArrow
Arrow	upArrow
Arrow	upDownArrow
Arrow	uturnArrow
Basic	blockArc
Basic	can
Basic	chevron
Basic	cube
Basic	decagon
Basic	diamond
Basic	dodecagon
Basic	donut
Basic	ellipse
Basic	funnel
Basic	gear6
Basic	gear9
Basic	heart
Basic	heptagon
Basic	hexagon
Basic	homePlate
Basic	lightningBolt
Basic	line
Basic	lineInv
Basic	moon
Basic	nonIsoscelesTrapezoid
Basic	noSmoking
Basic	octagon
Basic	parallelogram
Basic	pentagon
Basic	pie
Basic	pieWedge
Basic	plaque
Basic	rect
Basic	round1Rect
Basic	round2DiagRect
Basic	round2SameRect
Basic	roundRect
Basic	rtTriangle
Basic	smileyFace
Basic	snip1Rect
Basic	snip2DiagRect
Basic	snip2SameRect
Basic	snipRoundRect
Basic	star10
Basic	star12
Basic	star16
Basic	star24
Basic	star32
Basic	star4
Basic	star5
Basic	star6
Basic	star7
Basic	star8
Basic	sun
Basic	teardrop
Basic	trapezoid
Basic	triangle
Callout	accentBorderCallout1
Callout	accentBorderCallout2
Callout	accentBorderCallout3
Callout	accentCallout1
Callout	accentCallout2
Callout	accentCallout3
Callout	borderCallout1
Callout	borderCallout2
Callout	borderCallout3
Callout	callout1
Callout	callout2
Callout	callout3
Callout	cloudCallout
Callout	downArrowCallout
Callout	leftArrowCallout
Callout	leftRightArrowCallout
Callout	quadArrowCallout
Callout	rightArrowCallout
Callout	upArrowCallout
Callout	upDownArrowCallout
Callout	wedgeEllipseCallout
Callout	wedgeRectCallout
Callout	wedgeRoundRectCallout
Chart	chartPlus
Chart	chartStar
Chart	chartX
Connector	bentConnector2
Connector	bentConnector3
Connector	bentConnector4
Connector	bentConnector5
Connector	curvedConnector2
Connector	curvedConnector3
Connector	curvedConnector4
Connector	curvedConnector5
Connector	straightConnector1
FlowChart	flowChartAlternateProcess
FlowChart	flowChartCollate
FlowChart	flowChartConnector
FlowChart	flowChartDecision
FlowChart	flowChartDelay
FlowChart	flowChartDisplay
FlowChart	flowChartDocument
FlowChart	flowChartExtract
FlowChart	flowChartInputOutput
FlowChart	flowChartInternalStorage
FlowChart	flowChartMagneticDisk
FlowChart	flowChartMagneticDrum
FlowChart	flowChartMagneticTape
FlowChart	flowChartManualInput
FlowChart	flowChartManualOperation
FlowChart	flowChartMerge
FlowChart	flowChartMultidocument
FlowChart	flowChartOfflineStorage
FlowChart	flowChartOffpageConnector
FlowChart	flowChartOnlineStorage
FlowChart	flowChartOr
FlowChart	flowChartPredefinedProcess
FlowChart	flowChartPreparation
FlowChart	flowChartProcess
FlowChart	flowChartPunchedCard
FlowChart	flowChartPunchedTape
FlowChart	flowChartSort
FlowChart	flowChartSummingJunction
FlowChart	flowChartTerminator
Math	mathDivide
Math	mathEqual
Math	mathMinus
Math	mathMultiply
Math	mathNotEqual
Math	mathPlus
Star_Banner	arc
Star_Banner	bevel
Star_Banner	bracePair
Star_Banner	bracketPair
Star_Banner	chord
Star_Banner	cloud
Star_Banner	corner
Star_Banner	diagStripe
Star_Banner	doubleWave
Star_Banner	ellipseRibbon
Star_Banner	ellipseRibbon2
Star_Banner	foldedCorner
Star_Banner	frame
Star_Banner	halfFrame
Star_Banner	horizontalScroll
Star_Banner	irregularSeal1
Star_Banner	irregularSeal2
Star_Banner	leftBrace
Star_Banner	leftBracket
Star_Banner	leftRightRibbon
Star_Banner	plus
Star_Banner	ribbon
Star_Banner	ribbon2
Star_Banner	rightBrace
Star_Banner	rightBracket
Star_Banner	verticalScroll
Star_Banner	wave
Tabs	cornerTabs
Tabs	plaqueTabs
Tabs	squareTabs
EOS
  end

  def test_stats
    @xlsx = 'stats.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet('Test data')

    # Set the column width for column 1
    worksheet.set_column(0, 0, 20)

    # Create a format for the headings
    format = workbook.add_format
    format.set_bold

    # Write the sample data
    worksheet.write(0, 0, 'Sample', format)
    worksheet.write(0, 1, 1)
    worksheet.write(0, 2, 2)
    worksheet.write(0, 3, 3)
    worksheet.write(0, 4, 4)
    worksheet.write(0, 5, 5)
    worksheet.write(0, 6, 6)
    worksheet.write(0, 7, 7)
    worksheet.write(0, 8, 8)

    worksheet.write(1, 0, 'Length', format)
    worksheet.write(1, 1, 25.4)
    worksheet.write(1, 2, 25.4)
    worksheet.write(1, 3, 24.8)
    worksheet.write(1, 4, 25.0)
    worksheet.write(1, 5, 25.3)
    worksheet.write(1, 6, 24.9)
    worksheet.write(1, 7, 25.2)
    worksheet.write(1, 8, 24.8)

    # Write some statistical functions
    worksheet.write(4, 0, 'Count', format)
    worksheet.write(4, 1, '=COUNT(B1:I1)')

    worksheet.write(5, 0, 'Sum', format)
    worksheet.write(5, 1, '=SUM(B2:I2)')

    worksheet.write(6, 0, 'Average', format)
    worksheet.write(6, 1, '=AVERAGE(B2:I2)')

    worksheet.write(7, 0, 'Min', format)
    worksheet.write(7, 1, '=MIN(B2:I2)')

    worksheet.write(8, 0, 'Max', format)
    worksheet.write(8, 1, '=MAX(B2:I2)')

    worksheet.write(9, 0, 'Standard Deviation', format)
    worksheet.write(9, 1, '=STDEV(B2:I2)')

    worksheet.write(10, 0, 'Kurtosis', format)
    worksheet.write(10, 1, '=KURT(B2:I2)')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_stats_ext
    @xlsx = 'stats_ext.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet('Test results')
    worksheet2 = workbook.add_worksheet('Data')

    # Set the column width for column 1
    worksheet1.set_column(0, 0, 20)

    # Create a format for the headings
    headings = workbook.add_format
    headings.set_bold

    # Create a numerical format
    numformat = workbook.add_format
    numformat.set_num_format('0.00')

    # Write some statistical functions
    worksheet1.write(0, 0, 'Count', headings)
    worksheet1.write(0, 1, '=COUNT(Data!B2:B9)')

    worksheet1.write(1, 0, 'Sum', headings)
    worksheet1.write(1, 1, '=SUM(Data!B2:B9)')

    worksheet1.write(2, 0, 'Average', headings)
    worksheet1.write(2, 1, '=AVERAGE(Data!B2:B9)')

    worksheet1.write(3, 0, 'Min', headings)
    worksheet1.write(3, 1, '=MIN(Data!B2:B9)')

    worksheet1.write(4, 0, 'Max', headings)
    worksheet1.write(4, 1, '=MAX(Data!B2:B9)')

    worksheet1.write(5, 0, 'Standard Deviation', headings)
    worksheet1.write(5, 1, '=STDEV(Data!B2:B9)')

    worksheet1.write(6, 0, 'Kurtosis', headings)
    worksheet1.write(6, 1, '=KURT(Data!B2:B9)')

    # Write the sample data
    worksheet2.write(0, 0, 'Sample', headings)
    worksheet2.write(1, 0, 1)
    worksheet2.write(2, 0, 2)
    worksheet2.write(3, 0, 3)
    worksheet2.write(4, 0, 4)
    worksheet2.write(5, 0, 5)
    worksheet2.write(6, 0, 6)
    worksheet2.write(7, 0, 7)
    worksheet2.write(8, 0, 8)

    worksheet2.write(0, 1, 'Length', headings)
    worksheet2.write(1, 1, 25.4, numformat)
    worksheet2.write(2, 1, 25.4, numformat)
    worksheet2.write(3, 1, 24.8, numformat)
    worksheet2.write(4, 1, 25.0, numformat)
    worksheet2.write(5, 1, 25.3, numformat)
    worksheet2.write(6, 1, 24.9, numformat)
    worksheet2.write(7, 1, 25.2, numformat)
    worksheet2.write(8, 1, 24.8, numformat)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_stocks
    @xlsx = 'stocks.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width for columns 1, 2, 3 and 4
    worksheet.set_column(0, 3, 15)

    # Create a format for the column headings
    header = workbook.add_format
    header.set_bold
    header.set_size(12)
    header.set_color('blue')

    # Create a format for the stock price
    f_price = workbook.add_format
    f_price.set_align('left')
    f_price.set_num_format('$0.00')

    # Create a format for the stock volume
    f_volume = workbook.add_format
    f_volume.set_align('left')
    f_volume.set_num_format('#,##0')

    # Create a format for the price change. This is an example of a conditional
    # format. The number is formatted as a percentage. If it is positive it is
    # formatted in green, if it is negative it is formatted in red and if it is
    # zero it is formatted as the default font colour (in this case black).
    # Note: the [Green] format produces an unappealing lime green. Try
    # [Color 10] instead for a dark green.
    #
    f_change = workbook.add_format
    f_change.set_align('left')
    f_change.set_num_format('[Green]0.0%;[Red]-0.0%;0.0%')

    # Write out the data
    worksheet.write(0, 0, 'Company', header)
    worksheet.write(0, 1, 'Price',   header)
    worksheet.write(0, 2, 'Volume',  header)
    worksheet.write(0, 3, 'Change',  header)

    worksheet.write(1, 0, 'Damage Inc.')
    worksheet.write(1, 1, 30.25, f_price)       # $30.25
    worksheet.write(1, 2, 1234567, f_volume)    # 1,234,567
    worksheet.write(1, 3, 0.085, f_change)      # 8.5% in green

    worksheet.write(2, 0, 'Dump Corp.')
    worksheet.write(2, 1, 1.56, f_price)        # $1.56
    worksheet.write(2, 2, 7564, f_volume)       # 7,564
    worksheet.write(2, 3, -0.015, f_change)     # -1.5% in red

    worksheet.write(3, 0, 'Rev Ltd.')
    worksheet.write(3, 1, 0.13, f_price)        # $0.13
    worksheet.write(3, 2, 321, f_volume)        # 321
    worksheet.write(3, 3, 0, f_change)          # 0 in the font color (black)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_tab_colors
    @xlsx = 'tab_colors.xlsx'
    workbook = WriteXLSX.new(@io)

    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet

    # Worksheet1 will have the default tab colour.
    worksheet2.tab_color = 'red'
    worksheet3.tab_color = 'green'
    worksheet4.tab_color = 0x35    # Orange

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_tables
    @xlsx = 'tables.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet1  = workbook.add_worksheet
    worksheet2  = workbook.add_worksheet
    worksheet3  = workbook.add_worksheet
    worksheet4  = workbook.add_worksheet
    worksheet5  = workbook.add_worksheet
    worksheet6  = workbook.add_worksheet
    worksheet7  = workbook.add_worksheet
    worksheet8  = workbook.add_worksheet
    worksheet9  = workbook.add_worksheet
    worksheet10 = workbook.add_worksheet
    worksheet11 = workbook.add_worksheet
    worksheet12 = workbook.add_worksheet
    worksheet13 = workbook.add_worksheet

    currency_format = workbook.add_format(:num_format => '$#,##0')

    # Some sample data for the table.
    data = [
      ['Apples',  10000, 5000, 8000, 6000],
      ['Pears',   2000,  3000, 4000, 5000],
      ['Bananas', 6000,  6000, 6500, 6000],
      ['Oranges', 500,   300,  200,  700]
    ]

    ###############################################################################
    #
    # Example 1.
    #
    caption = 'Default table with no data.'

    # Set the columns widths.
    worksheet1.set_column('B:G', 12)

    # Write the caption.
    worksheet1.write('B1', caption)

    # Add a table to the worksheet.
    worksheet1.add_table('B3:F7')

    ###############################################################################
    #
    # Example 2.
    #
    caption = 'Default table with data.'

    # Set the columns widths.
    worksheet2.set_column('B:G', 12)

    # Write the caption.
    worksheet2.write('B1', caption)

    # Add a table to the worksheet.
    worksheet2.add_table('B3:F7', { :data => data })

    ###############################################################################
    #
    # Example 3.
    #
    caption = 'Table without default autofilter.'

    # Set the columns widths.
    worksheet3.set_column('B:G', 12)

    # Write the caption.
    worksheet3.write('B1', caption)

    # Add a table to the worksheet.
    worksheet3.add_table('B3:F7', { :autofilter => 0 })

    # Table data can also be written separately, as an array or individual cells.
    worksheet3.write_col('B4', data)

    ###############################################################################
    #
    # Example 4.
    #
    caption = 'Table without default header row.'

    # Set the columns widths.
    worksheet4.set_column('B:G', 12)

    # Write the caption.
    worksheet4.write('B1', caption)

    # Add a table to the worksheet.
    worksheet4.add_table('B4:F7', { :header_row => 0 })

    # Table data can also be written separately, as an array or individual cells.
    worksheet4.write_col('B4', data)

    ###############################################################################
    #
    # Example 5.
    #
    caption = 'Default table with "First Column" and "Last Column" options.'

    # Set the columns widths.
    worksheet5.set_column('B:G', 12)

    # Write the caption.
    worksheet5.write('B1', caption)

    # Add a table to the worksheet.
    worksheet5.add_table('B3:F7', { :first_column => 1, :last_column => 1 })

    # Table data can also be written separately, as an array or individual cells.
    worksheet5.write_col('B4', data)

    ###############################################################################
    #
    # Example 6.
    #
    caption = 'Table with banded columns but without default banded rows.'

    # Set the columns widths.
    worksheet6.set_column('B:G', 12)

    # Write the caption.
    worksheet6.write('B1', caption)

    # Add a table to the worksheet.
    worksheet6.add_table('B3:F7', { :banded_rows => 0, :banded_columns => 1 })

    # Table data can also be written separately, as an array or individual cells.
    worksheet6.write_col('B4', data)

    ###############################################################################
    #
    # Example 7.
    #
    caption = 'Table with user defined column headers'

    # Set the columns widths.
    worksheet7.set_column('B:G', 12)

    # Write the caption.
    worksheet7.write('B1', caption)

    # Add a table to the worksheet.
    worksheet7.add_table(
      'B3:F7',
      {
        :data    => data,
        :columns => [
          { :header => 'Product' },
          { :header => 'Quarter 1' },
          { :header => 'Quarter 2' },
          { :header => 'Quarter 3' },
          { :header => 'Quarter 4' }
        ]
      }
    )

    ###############################################################################
    #
    # Example 8.
    #
    caption = 'Table with user defined column headers'

    # Set the columns widths.
    worksheet8.set_column('B:G', 12)

    # Write the caption.
    worksheet8.write('B1', caption)

    # Add a table to the worksheet.
    worksheet8.add_table(
      'B3:G7',
      {
        :data    => data,
        :columns => [
          { :header => 'Product' },
          { :header => 'Quarter 1' },
          { :header => 'Quarter 2' },
          { :header => 'Quarter 3' },
          { :header => 'Quarter 4' },
          {
            :header  => 'Year',
            :formula => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])'
          }
        ]
      }
    )

    ###############################################################################
    #
    # Example 9.
    #
    caption = 'Table with totals row (but no caption or totals).'

    # Set the columns widths.
    worksheet9.set_column('B:G', 12)

    # Write the caption.
    worksheet9.write('B1', caption)

    # Add a table to the worksheet.
    worksheet9.add_table(
      'B3:G8',
      {
        :data      => data,
        :total_row => 1,
        :columns   => [
          { :header => 'Product' },
          { :header => 'Quarter 1' },
          { :header => 'Quarter 2' },
          { :header => 'Quarter 3' },
          { :header => 'Quarter 4' },
          {
            :header  => 'Year',
            :formula => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])'
          }
        ]
      }
    )

    ###############################################################################
    #
    # Example 10.
    #
    caption = 'Table with totals row with user captions and functions.'

    # Set the columns widths.
    worksheet10.set_column('B:G', 12)

    # Write the caption.
    worksheet10.write('B1', caption)

    # Add a table to the worksheet.
    worksheet10.add_table(
      'B3:G8',
      {
        :data      => data,
        :total_row => 1,
        :columns   => [
          { :header => 'Product',   :total_string   => 'Totals' },
          { :header => 'Quarter 1', :total_function => 'sum' },
          { :header => 'Quarter 2', :total_function => 'sum' },
          { :header => 'Quarter 3', :total_function => 'sum' },
          { :header => 'Quarter 4', :total_function => 'sum' },
          {
            :header         => 'Year',
            :formula        => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])',
            :total_function => 'sum'
          }
        ]
      }
    )

    ###############################################################################
    #
    # Example 11.
    #
    caption = 'Table with alternative Excel style.'

    # Set the columns widths.
    worksheet11.set_column('B:G', 12)

    # Write the caption.
    worksheet11.write('B1', caption)

    # Add a table to the worksheet.
    worksheet11.add_table(
      'B3:G8',
      {
        :data      => data,
        :style     => 'Table Style Light 11',
        :total_row => 1,
        :columns   => [
          { :header => 'Product',   :total_string   => 'Totals' },
          { :header => 'Quarter 1', :total_function => 'sum' },
          { :header => 'Quarter 2', :total_function => 'sum' },
          { :header => 'Quarter 3', :total_function => 'sum' },
          { :header => 'Quarter 4', :total_function => 'sum' },
          {
            :header         => 'Year',
            :formula        => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])',
            :total_function => 'sum'
          }
        ]
      }
    )

    ###############################################################################
    #
    # Example 12.
    #
    caption = 'Table with no Excel style.'

    # Set the columns widths.
    worksheet12.set_column('B:G', 12)

    # Write the caption.
    worksheet12.write('B1', caption)

    # Add a table to the worksheet.
    worksheet12.add_table(
      'B3:G8',
      {
        :data      => data,
        :style     => 'None',
        :total_row => 1,
        :columns   => [
          { :header => 'Product',   :total_string   => 'Totals' },
          { :header => 'Quarter 1', :total_function => 'sum' },
          { :header => 'Quarter 2', :total_function => 'sum' },
          { :header => 'Quarter 3', :total_function => 'sum' },
          { :header => 'Quarter 4', :total_function => 'sum' },
          {
            :header         => 'Year',
            :formula        => '=SUM(Table12[@[Quarter 1]:[Quarter 4]])',
            :total_function => 'sum'
          }
        ]
      }
    )

    ###############################################################################
    #
    # Example 13.
    #
    caption = 'Table with column formats.'

    # Set the columns widths.
    worksheet13.set_column('B:G', 12)

    # Write the caption.
    worksheet13.write('B1', caption)

    # Add a table to the worksheet.
    worksheet13.add_table(
      'B3:G8',
      {
        :data      => data,
        :total_row => 1,
        :columns   => [
          { :header => 'Product', :total_string => 'Totals' },
          {
            :header         => 'Quarter 1',
            :total_function => 'sum',
            :format         => currency_format
          },
          {
            :header         => 'Quarter 2',
            :total_function => 'sum',
            :format         => currency_format
          },
          {
            :header         => 'Quarter 3',
            :total_function => 'sum',
            :format         => currency_format
          },
          {
            :header         => 'Quarter 4',
            :total_function => 'sum',
            :format         => currency_format
          },
          {
            :header         => 'Year',
            :formula        => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])',
            :total_function => 'sum',
            :format         => currency_format
          }
        ]
      }
    )

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_sparklines1
    @xlsx = 'sparklines1.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Some sample data to plot.
    data = [
      [-2, 2,  3,  -1, 0],
      [30, 20, 33, 20, 15],
      [1,  -1, -1, 1,  -1]
    ]

    # Write the sample data to the worksheet.
    worksheet.write_col('A1', data)

    # Add a line sparkline (the default) with markers.
    worksheet.add_sparkline(
      {
        :location => 'F1',
        :range    => 'Sheet1!A1:E1',
        :markers  => 1
      }
    )

    # Add a column sparkline with non-default style.
    worksheet.add_sparkline(
      {
        :location => 'F2',
        :range    => 'Sheet1!A2:E2',
        :type     => 'column',
        :style    => 12
      }
    )

    # Add a win/loss sparkline with negative values highlighted.
    worksheet.add_sparkline(
      {
        :location        => 'F3',
        :range           => 'Sheet1!A3:E3',
        :type            => 'win_loss',
        :negative_points => 1
      }
    )

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_sparklines2
    @xlsx = 'sparklines2.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    bold       = workbook.add_format(:bold => 1)
    row = 1

    # Set the columns widths to make the output clearer.
    worksheet1.set_column('A:A', 14)
    worksheet1.set_column('B:B', 50)
    worksheet1.zoom = 150

    # Headings.
    worksheet1.write('A1', 'Sparkline',   bold)
    worksheet1.write('B1', 'Description', bold)

    ##########################################################################
    #
    str = 'A default "line" sparkline.'

    worksheet1.add_sparkline(
      {
        :location => 'A2',
        :range    => 'Sheet2!A1:J1'
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'A default "column" sparkline.'

    worksheet1.add_sparkline(
      {
        :location => 'A3',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column'
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'A default "win/loss" sparkline.'

    worksheet1.add_sparkline(
      {
        :location => 'A4',
        :range    => 'Sheet2!A3:J3',
        :type     => 'win_loss'
      }
    )

    worksheet1.write(row, 1, str)
    row += 2

    ##########################################################################
    #
    str = 'Line with markers.'

    worksheet1.add_sparkline(
      {
        :location => 'A6',
        :range    => 'Sheet2!A1:J1',
        :markers  => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Line with high and low points.'

    worksheet1.add_sparkline(
      {
        :location   => 'A7',
        :range      => 'Sheet2!A1:J1',
        :high_point => 1,
        :low_point  => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Line with first and last point markers.'

    worksheet1.add_sparkline(
      {
        :location    => 'A8',
        :range       => 'Sheet2!A1:J1',
        :first_point => 1,
        :last_point  => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Line with negative point markers.'

    worksheet1.add_sparkline(
      {
        :location        => 'A9',
        :range           => 'Sheet2!A1:J1',
        :negative_points => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Line with axis.'

    worksheet1.add_sparkline(
      {
        :location => 'A10',
        :range    => 'Sheet2!A1:J1',
        :axis     => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 2

    ##########################################################################
    #
    str = 'Column with default style (1).'

    worksheet1.add_sparkline(
      {
        :location => 'A12',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column'
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Column with style 2.'

    worksheet1.add_sparkline(
      {
        :location => 'A13',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column',
        :style    => 2
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Column with style 3.'

    worksheet1.add_sparkline(
      {
        :location => 'A14',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column',
        :style    => 3
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Column with style 4.'

    worksheet1.add_sparkline(
      {
        :location => 'A15',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column',
        :style    => 4
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Column with style 5.'

    worksheet1.add_sparkline(
      {
        :location => 'A16',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column',
        :style    => 5
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Column with style 6.'

    worksheet1.add_sparkline(
      {
        :location => 'A17',
        :range    => 'Sheet2!A2:J2',
        :type     => 'column',
        :style    => 6
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Column with a user defined colour.'

    worksheet1.add_sparkline(
      {
        :location     => 'A18',
        :range        => 'Sheet2!A2:J2',
        :type         => 'column',
        :series_color => '#E965E0'
      }
    )

    worksheet1.write(row, 1, str)
    row += 2

    ##########################################################################
    #
    str = 'A win/loss sparkline.'

    worksheet1.add_sparkline(
      {
        :location => 'A20',
        :range    => 'Sheet2!A3:J3',
        :type     => 'win_loss'
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'A win/loss sparkline with negative points highlighted.'

    worksheet1.add_sparkline(
      {
        :location        => 'A21',
        :range           => 'Sheet2!A3:J3',
        :type            => 'win_loss',
        :negative_points => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 2

    ##########################################################################
    #
    str = 'A left to right column (the default).'

    worksheet1.add_sparkline(
      {
        :location => 'A23',
        :range    => 'Sheet2!A4:J4',
        :type     => 'column',
        :style    => 20
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'A right to left column.'

    worksheet1.add_sparkline(
      {
        :location => 'A24',
        :range    => 'Sheet2!A4:J4',
        :type     => 'column',
        :style    => 20,
        :reverse  => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    #
    str = 'Sparkline and text in one cell.'

    worksheet1.add_sparkline(
      {
        :location => 'A25',
        :range    => 'Sheet2!A4:J4',
        :type     => 'column',
        :style    => 20
      }
    )

    worksheet1.write(row,   0, 'Growth')
    worksheet1.write(row, 1, str)
    row += 2

    ##########################################################################
    #
    str = 'A grouped sparkline. Changes are applied to all three.'

    worksheet1.add_sparkline(
      {
        :location => %w[A27 A28 A29],
        :range    => ['Sheet2!A5:J5', 'Sheet2!A6:J6', 'Sheet2!A7:J7'],
        :markers  => 1
      }
    )

    worksheet1.write(row, 1, str)
    row += 1

    ##########################################################################
    # Create a second worksheet with data to plot.
    #

    worksheet2.set_column('A:J', 11)

    data = [
      # Simple line data.
      [-2, 2, 3, -1, 0, -2, 3, 2, 1, 0],

      # Simple column data.
      [30, 20, 33, 20, 15, 5, 5, 15, 10, 15],

      # Simple win/loss data.
      [1, 1, -1, -1, 1, -1, 1, 1, 1, -1],

      # Unbalanced histogram.
      [5, 6, 7, 10, 15, 20, 30, 50, 70, 100],

      # Data for the grouped sparkline example.
      [-2, 2,  3, -1, 0, -2, 3, 2, 1, 0],
      [3,  -1, 0, -2, 3, 2,  1, 0, 2, 1],
      [0,  -2, 3, 2,  1, 0,  1, 2, 3, 1]
    ]

    # Write the sample data to the worksheet.
    worksheet2.write_col('A1', data)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_add_vba_project
    @xlsx = 'add_vba_project.xlsm'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('A:A', 50)

    # Add the VBA project binary.
    workbook.add_vba_project(File.join(@test_dir, 'vbaProject.bin'))

    # Show text for the end user.
    worksheet.write('A1', 'Run the SampleMacro embedded in this file.')
    worksheet.write('A2', 'You may have to turn on the Excel Developer option first.')

    # Call a user defined function from the VBA project.
    worksheet.write('A6', 'Result from a user defined function:')
    worksheet.write('B6', '=MyFunction(7)')

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_hide_row_col
    @xlsx = 'hide_row_col.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Write some data
    worksheet.write('D1', 'Some hidden columns.')
    worksheet.write('A8', 'Some hidden rows.')

    # Hide all rows without data.
    worksheet.set_default_row(nil, 1)

    # Set emptys row that we do want to display. All other will be hidden.
    (1..6).each { |row| worksheet.set_row(row, 15) }

    # Hide a range of columns.
    worksheet.set_column('G:XFD', nil, nil, 1)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_data_table
    @xlsx = 'chart_data_table.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2, 3, 4, 5, 6, 7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new column chart with a data table.
    chart1 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the first series.
    chart1.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart1.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart1.set_title(:name => 'Chart with Data Table')
    chart1.set_x_axis(:name => 'Test number')
    chart1.set_y_axis(:name => 'Sample length (mm)')

    # Set a default data table on the X-Axis.
    chart1.set_table

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart1, 25, 10)

    #
    # Create a second charat.
    #
    chart2 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the first series.
    chart2.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ sheetname, row_start, row_end, col_start, col_end ].
    chart2.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => ['Sheet1', 1, 6, 0, 0],
      :values     => ['Sheet1', 1, 6, 2, 2]
    )

    # Add a chart title and some axis labels.
    chart2.set_title(:name => 'Data Table with legend keys')
    chart2.set_x_axis(:name => 'Test number')
    chart2.set_y_axis(:name => 'Sample length (mm)')

    # Set a default data table on the X-Axis with the legend keys shown.
    chart2.set_table(:show_keys => true)

    # Hide the chart legend since the keys are show on the data table.
    chart2.set_legend(:position => 'none')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D18', chart2, 25, 11)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_data_tools
    @xlsx = 'chart_data_tools.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Data 1', 'Data 2']
    data = [
      [2,  3,  4,  5,  6,  7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    #######################################################################
    #
    # Trendline example.
    #

    # Create a Line chart.
    chart1 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the first series with a polynomial trendline.
    chart1.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7',
      :trendline  => {
        :type  => 'polynomial',
        :order => 3
      }
    )

    # Configure the second series with a moving average trendline.
    chart1.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7',
      :trendline  => { :type => 'linear' }
    )

    # Add a chart title. and some axis labels.
    chart1.set_title(:name => 'Chart with Trendlines')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart1, 25, 10)

    #######################################################################
    #
    # Data Labels and Markers example.
    #

    # Create a Line chart.
    chart2 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the first series.
    chart2.add_series(
      :categories  => '=Sheet1!$A$2:$A$7',
      :values      => '=Sheet1!$B$2:$B$7',
      :data_labels => { :value => 1 },
      :marker      => { :type => 'automatic' }
    )

    # Configure the second series.
    chart2.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7'
    )

    # Add a chart title. and some axis labels.
    chart2.set_title(:name => 'Chart with Data Labels and Markers')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D18', chart2, 25, 10)

    #######################################################################
    #
    # Error Bars example.
    #

    # Create a Line chart.
    chart3 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the first series.
    chart3.add_series(
      :categories   => '=Sheet1!$A$2:$A$7',
      :values       => '=Sheet1!$B$2:$B$7',
      :y_error_bars => { :type => 'standard_error' }
    )

    # Configure the second series.
    chart3.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7'
    )

    # Add a chart title. and some axis labels.
    chart3.set_title(:name => 'Chart with Error Bars')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D34', chart3, 25, 10)

    #######################################################################
    #
    # Up-Down Bars example.
    #

    # Create a Line chart.
    chart4 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Add the Up-Down Bars.
    chart4.set_up_down_bars

    # Configure the first series.
    chart4.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure the second series.
    chart4.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7'
    )

    # Add a chart title. and some axis labels.
    chart4.set_title(:name => 'Chart with Up-Down Bars')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D50', chart4, 25, 10)

    #######################################################################
    #
    # High-Low Lines example.
    #

    # Create a Line chart.
    chart5 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Add the High-Low lines.
    chart5.set_high_low_lines

    # Configure the first series.
    chart5.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure the second series.
    chart5.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7'
    )

    # Add a chart title. and some axis labels.
    chart5.set_title(:name => 'Chart with High-Low Lines')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D66', chart5, 25, 10)

    #######################################################################
    #
    # Drop Lines example.
    #

    # Create a Line chart.
    chart6 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Add Drop Lines.
    chart6.set_drop_lines

    # Configure the first series.
    chart6.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Configure the second series.
    chart6.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7'
    )

    # Add a chart title. and some axis labels.
    chart6.set_title(:name => 'Chart with Drop Lines')

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D82', chart6, 25, 10)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_combined
    @xlsx = 'chart_combined.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
      [2,  3,  4,  5,  6,  7],
      [10, 40, 50, 20, 10, 50],
      [30, 60, 70, 50, 40, 30]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    #
    # In the first example we will create a combined column and line chart.
    # They will share the same X and Y axes.
    #

    # Create a new column chart. This will use this as the primary chart.
    column_chart1 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the data series for the primary chart.
    column_chart1.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Create a new column chart. This will use this as the secondary chart.
    line_chart1 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the data series for the secondary chart.
    line_chart1.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7'
    )

    # Combine the charts.
    column_chart1.combine(line_chart1)

    # Add a chart title and some axis labels. Note, this is done via the
    # primary chart.
    column_chart1.set_title(:name => 'Combined chart - same Y axis')
    column_chart1.set_x_axis(:name => 'Test number')
    column_chart1.set_y_axis(:name => 'Sample length (mm)')

    # Insert the chart into the worksheet
    worksheet.insert_chart('E2', column_chart1)

    #
    # In the second example we will create a similar combined column and line
    # chart except that the secondary chart will have a secondary Y axis.
    #

    # Create a new column chart. This will use this as the primary chart.
    column_chart2 = workbook.add_chart(:type => 'column', :embedded => 1)

    # Configure the data series for the primary chart.
    column_chart2.add_series(
      :name       => '=Sheet1!$B$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$B$2:$B$7'
    )

    # Create a new column chart. This will use this as the secondary chart.
    line_chart2 = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the data series for the secondary chart. We also set a
    # secondary Y axis via (y2_axis). This is the only difference between
    # this and the first example, apart from the axis label below.
    line_chart2.add_series(
      :name       => '=Sheet1!$C$1',
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7',
      :y2_axis    => 1
    )

    # Combine the charts.
    column_chart2.combine(line_chart2)

    # Add a chart title and some axis labels.
    column_chart2.set_title(:name => 'Combine chart - secondary Y axis')
    column_chart2.set_x_axis(:name => 'Test number')
    column_chart2.set_y_axis(:name => 'Sample length (mm)')
    column_chart2.set_y2_axis(:name => 'Target length (mm)')

    # Insert the chart into the worksheet
    worksheet.insert_chart('E18', column_chart2)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end

  def test_chart_pareto
    @xlsx = 'chart_pareto.xlsx'
    workbook = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Formats used in the workbook.
    bold           = workbook.add_format(:bold => 1)
    percent_format = workbook.add_format(:num_format => '0.0%')

    # Widen the columns for visibility.
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:C', 10)

    # Add the worksheet data that the charts will refer to.
    headings = %w[Reason Number Percentage]

    reasons = [
      'Traffic',   'Child care', 'Public Transport', 'Weather',
      'Overslept', 'Emergency'
    ]

    numbers  = [60,    40,  20,  15,    10, 5]
    percents = [0.44, 0.667, 0.8, 0.9, 0.967, 1]

    worksheet.write_row('A1', headings, bold)
    worksheet.write_col('A2', reasons)
    worksheet.write_col('B2', numbers)
    worksheet.write_col('C2', percents, percent_format)

    # Create a new column chart. This will be the primary chart.
    column_chart = workbook.add_chart(:type => 'column', :embedded => 1)

    # Add a series
    column_chart.add_series(
      :categories => 'Sheet1!$A$2:$A$7',
      :values     => 'Sheet1!$B$2:$B$7'
    )

    # Add a chart title.
    column_chart.set_title(:name => 'Reasons for lateness')

    # Turn off the chart legend.
    column_chart.set_legend(:position => 'none')

    # Set the title and scale of the Y axes. Note, the secondary axis is set from
    # the primary chart.
    column_chart.set_y_axis(
      :name => 'Respondents (number)',
      :min  => 0,
      :max  => 120
    )
    column_chart.set_y2_axis(:max => 1)

    # Create a new line chart. This will be the secondary chart.
    line_chart = workbook.add_chart(:type => 'line', :embedded => 1)

    # Add a series, on the secondary axis.
    line_chart.add_series(
      :categories => '=Sheet1!$A$2:$A$7',
      :values     => '=Sheet1!$C$2:$C$7',
      :marker     => { :type => 'automatic' },
      :y2_axis    => 1
    )

    # Combine the charts.
    column_chart.combine(line_chart)

    # Insert the chart into the worksheet.
    worksheet.insert_chart('F2', column_chart)

    workbook.close
    store_to_tempfile
    compare_xlsx(File.join(@perl_output, @xlsx), @tempfile.path)
  end
end
