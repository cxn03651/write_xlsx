# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'

class TestExampleMatch < Test::Unit::TestCase
  def setup
    setup_dir_var
    system("rm -rf #{@expected_dir}") if File.exist?(@expected_dir)
    system("rm -rf #{@result_dir}")   if File.exist?(@result_dir)
    raise "cannot create test working directory." if File.exist?(@expected_dir) || File.exist?(@result_dir)
    @obj = Writexlsx::Package::XMLWriterSimple.new
  end

  def test_a_simple
    xlsx = 'a_simple.xlsx'
    # Create a new workbook called simple.xls and add a worksheet
    workbook  = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_array_formula
    xlsx = 'array_formula.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet

    # Write some test data.
    worksheet.write('B1', [ [ 500, 10 ], [ 300, 15 ] ])
    worksheet.write('B5', [ [ 1, 2, 3 ], [ 20234, 21003, 10000 ] ])

    # Write an array formula that returns a single value
    worksheet.write('A1', '{=SUM(B1:C1*B2:C2)}')

    # Same as above but more verbose.
    worksheet.write_array_formula('A2:A2', '{=SUM(B1:C1*B2:C2)}')

    # Write an array formula that returns a range of values
    worksheet.write_array_formula('A5:A7', '{=TREND(C5:C7,B5:B7)}')

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_area
    xlsx = 'chart_area.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = [ 'Number', 'Batch 1', 'Batch 2' ]
    data = [
            [ 2, 3, 4, 5, 6, 7 ],
            [ 40, 40, 50, 30, 25, 50 ],
            [ 30, 25, 30, 10,  5, 10 ]
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
                     :categories => [ 'Sheet1', 1, 6, 0, 0 ],
                     :values     => [ 'Sheet1', 1, 6, 2, 2 ]
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_bar
    xlsx = 'chart_bar.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = [ 'Number', 'Batch 1', 'Batch 2' ]
    data = [
        [ 2, 3, 4, 5, 6, 7 ],
        [ 10, 40, 50, 20, 10, 50 ],
        [ 30, 60, 70, 50, 40, 30 ]
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
        :categories => [ 'Sheet1', 1, 6, 0, 0 ],
        :values     => [ 'Sheet1', 1, 6, 2, 2 ]
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_column
    xlsx = 'chart_column.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = [ 'Number', 'Batch 1', 'Batch 2' ]
    data = [
        [ 2, 3, 4, 5, 6, 7 ],
        [ 10, 40, 50, 20, 10, 50 ],
        [ 30, 60, 70, 50, 40, 30 ]
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
        :categories => [ 'Sheet1', 1, 6, 0, 0 ],
        :values     => [ 'Sheet1', 1, 6, 2, 2 ]
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_line
    xlsx = 'chart_line.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = [ 'Number', 'Batch 1', 'Batch 2' ]
    data = [
        [ 2, 3, 4, 5, 6, 7 ],
        [ 10, 40, 50, 20, 10, 50 ],
        [ 30, 60, 70, 50, 40, 30 ]
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
        :categories => [ 'Sheet1', 1, 6, 0, 0 ],
        :values     => [ 'Sheet1', 1, 6, 2, 2 ]
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_pie
    xlsx = 'chart_pie.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = [ 'Category', 'Values' ]
    data = [
        [ 'Apple', 'Cherry', 'Pecan' ],
        [ 60,       30,       10     ]
    ]

    worksheet.write('A1', headings, bold)
    worksheet.write('A2', data)

    # Create a new chart object. In this case an embedded chart.
    chart = workbook.add_chart(:type => 'pie', :embedded => 1)

    # Configure the series. Note the use of the array ref to define ranges:
    # [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
    chart.add_series(
        :name       => 'Pie sales data',
        :categories => [ 'Sheet1', 1, 3, 0, 0 ],
        :values     => [ 'Sheet1', 1, 3, 1, 1 ]
    )

    # Add a title.
    chart.set_title(:name => 'Popular Pie Types')

    # Set an Excel chart style. Blue colors with white outline and shadow.
    chart.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C2', chart, 25, 10)

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_scatter
    xlsx = 'chart_scatter.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Add the worksheet data that the charts will refer to.
    headings = [ 'Number', 'Batch 1', 'Batch 2' ]
    data = [
        [ 2, 3, 4, 5, 6, 7 ],
        [ 10, 40, 50, 20, 10, 50 ],
        [ 30, 60, 70, 50, 40, 30 ]
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
        :categories => [ 'Sheet1', 1, 6, 0, 0 ],
        :values     => [ 'Sheet1', 1, 6, 2, 2 ]
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_chart_stock
    xlsx = 'chart_stock.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet   = workbook.add_worksheet
    bold        = workbook.add_format(:bold => 1)
    date_format = workbook.add_format(:num_format => 'dd/mm/yyyy')
    chart       = workbook.add_chart(:type => 'stock', :embedded => 1)


    # Add the worksheet data that the charts will refer to.
    headings = [ 'Date', 'High', 'Low', 'Close' ]
    data = [
        [ '2007-01-01T', '2007-01-02T', '2007-01-03T', '2007-01-04T', '2007-01-05T' ],
        [ 27.2,  25.03, 19.05, 20.34, 18.5 ],
        [ 23.49, 19.55, 15.12, 17.84, 16.34 ],
        [ 25.45, 23.05, 17.32, 20.45, 17.34 ]
    ]

    worksheet.write('A1', headings, bold)

    (0 .. 4).each do |row|
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_conditional_format
    xlsx = 'conditional_format.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet1 = workbook.add_worksheet

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
            [ 90, 80,  50, 10,  20,  90,  40, 90,  30,  40 ],
            [ 20, 10,  90, 100, 30,  60,  70, 60,  50,  90 ],
            [ 10, 50,  60, 50,  20,  50,  80, 30,  40,  60 ],
            [ 10, 90,  20, 40,  10,  40,  50, 70,  90,  50 ],
            [ 70, 100, 10, 90,  10,  10,  20, 100, 100, 40 ],
            [ 20, 60,  10, 100, 30,  10,  20, 60,  100, 10 ],
            [ 10, 60,  10, 80,  100, 80,  30, 30,  70,  40 ],
            [ 30, 90,  60, 10,  10,  100, 40, 40,  30,  40 ],
            [ 80, 90,  10, 20,  20,  50,  80, 20,  60,  90 ],
            [ 60, 80,  30, 30,  10,  50,  80, 60,  50,  30 ]
           ]


    # This example below highlights cells that have a value greater than or
    # equal to 50 in red and cells below that value in green.

    caption = 'Cells with values >= 50 are in light red. ' +
      'Values < 50 are in light green'

    # Write the data.
    worksheet1.write('A1', caption)
    worksheet1.write_col('B3', data)

    # Write a conditional format over a range.
    worksheet1.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'cell',
                                        :format   => format1,
                                        :criteria => '>=',
                                        :value    => 50
                                      }
                                      )

    # Write another conditional format over the same range.
    worksheet1.conditional_formatting('B3:K12',
                                      {
                                        :type     => 'cell',
                                        :format   => format2,
                                        :criteria => '<',
                                        :value    => 50
                                      }
                                      )

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_data_validate
    xlsx = 'data_validate.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet

    # Add a format for the header cells.
    header_format = workbook.add_format(
                                        :border      => 1,
                                        :bg_color    => 43,
                                        :bold        => 1,
                                        :text_wrap   => 1,
                                        :valign      => 'vcenter',
                                        :indent      => 1
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
                                :validate        => 'integer',
                                :criteria        => 'between',
                                :minimum         => 1,
                                :maximum         => 10
                              })


    #
    # Example 2. Limiting input to an integer outside a fixed range.
    #
    txt = 'Enter an integer that is not between 1 and 10 (using cell references)'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'integer',
                                :criteria        => 'not between',
                                :minimum         => '=E3',
                                :maximum         => '=F3'
                              })


    #
    # Example 3. Limiting input to an integer greater than a fixed value.
    #
    txt = 'Enter an integer greater than 0'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'integer',
                                :criteria        => '>',
                                :value           => 0
                              })


    #
    # Example 4. Limiting input to an integer less than a fixed value.
    #
    txt = 'Enter an integer less than 10'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'integer',
                                :criteria        => '<',
                                :value           => 10
                              })


    #
    # Example 5. Limiting input to a decimal in a fixed range.
    #
    txt = 'Enter a decimal between 0.1 and 0.5'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'decimal',
                                :criteria        => 'between',
                                :minimum         => 0.1,
                                :maximum         => 0.5
                              })


    #
    # Example 6. Limiting input to a value in a dropdown list.
    #
    txt = 'Select a value from a drop down list'
    row += 2
    bp=1
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'list',
                                :source          => ['open', 'high', 'close']
                              })


    #
    # Example 6. Limiting input to a value in a dropdown list.
    #
    txt = 'Select a value from a drop down list (using a cell range)'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'list',
                                :source          => '=$E$4:$G$4'
                              })


    #
    # Example 7. Limiting input to a date in a fixed range.
    #
    txt = 'Enter a date between 1/1/2008 and 12/12/2008'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'date',
                                :criteria        => 'between',
                                :minimum         => '2008-01-01T',
                                :maximum         => '2008-12-12T'
                              })


    #
    # Example 8. Limiting input to a time in a fixed range.
    #
    txt = 'Enter a time between 6:00 and 12:00'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'time',
                                :criteria        => 'between',
                                :minimum         => 'T06:00',
                                :maximum         => 'T12:00'
                              })


    #
    # Example 9. Limiting input to a string greater than a fixed length.
    #
    txt = 'Enter a string longer than 3 characters'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'length',
                                :criteria        => '>',
                                :value           => 3
                              })


    #
    # Example 10. Limiting input based on a formula.
    #
    txt = 'Enter a value if the following is true "=AND(F5=50,G5=60)"'
    row += 2

    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
                              {
                                :validate        => 'custom',
                                :value           => '=AND(F5=50,G5=60)'
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_defined_name
    xlsx = 'defined_name.xlsx'
    workbook   = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_diag_border
    xlsx = 'diag_border.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet()


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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_headers
    xlsx = 'headers.xlsx'
    workbook = WriteXLSX.new(xlsx)
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
    header3    = %q(&C&"Courier New,Bold"Hello &"Arial,Italic"World)
    footer3    = %q(&C&"Symbol"e&"Arial" = mc&X2)

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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_hide_sheet
    xlsx = 'hide_sheet.xlsx'
    workbook   = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_hyperlink
    xlsx = 'hyperlink.xlsx'
    # Create a new workbook and add a worksheet
    workbook = WriteXLSX.new(xlsx)

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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_indent
    xlsx = 'indent.xlsx'
    workbook = WriteXLSX.new(xlsx)

    worksheet = workbook.add_worksheet
    indent1   = workbook.add_format(:indent => 1)
    indent2   = workbook.add_format(:indent => 2)

    worksheet.set_column('A:A', 40)

    worksheet.write('A1', "This text is indented 1 level",  indent1)
    worksheet.write('A2', "This text is indented 2 levels", indent2)

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_outline
    xlsx = 'outline.xlsx'
    # Create a new workbook and add some worksheets
    workbook   = WriteXLSX.new(xlsx)
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
            [ 'Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', ' Total' ],
            [ 'North', 50,    20,    15,    25,    65,    80,    '=SUM(B2:G2)' ],
            [ 'South', 10,    20,    30,    50,    50,    50,    '=SUM(B3:G3)' ],
            [ 'East',  45,    75,    50,    15,    75,    100,   '=SUM(B4:G4)' ],
            [ 'West',  15,    15,    55,    35,    20,    50,    '=SUM(B5:G5)' ],
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge1
    xlsx = 'merge1.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge2
    # Create a new workbook and add a worksheet
    xlsx = 'merge2.xlsx'
    workbook  = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge3
    xlsx = 'merge2.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge3
    xlsx = 'merge3.xlsx'

    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new('merge3.xlsx')
    worksheet = workbook.add_worksheet()

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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge4
    xlsx = 'merge4.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    (1 .. 11).each { |i| worksheet.set_row(i, 30) }
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

    worksheet.merge_range('B11:D12', 'Justified: ' << 'so on and ' * 18, format4)

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge5
    xlsx = 'merge5.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    (3 .. 8).each { |row|  worksheet.set_row(row, 36 ) }
    [1, 3, 5].each { |col| worksheet.set_column( col, col, 15 ) }

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

    worksheet.merge_range( 'B4:B9', 'Rotation 270', format1 )

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

    worksheet.merge_range( 'D4:D9', 'Rotation 90°', format2 )

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

    worksheet.merge_range( 'F4:F9', 'Rotation -90°', format3 )

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_merge6
    xlsx = 'merge6.xlsx'
    # Create a new workbook and add a worksheet
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    (2 .. 9).each { |row| worksheet.set_row(row, 36) }
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_comments1
    xlsx = 'comments1.xlsx'
    workbook  = WriteXLSX.new(xlsx)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Hello')
    worksheet.write_comment('A1', 'This is a comment')

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_comments2
    xlsx = 'comments2.xlsx'
    workbook  = WriteXLSX.new(xlsx)

    text_wrap  = workbook.add_format( :text_wrap => 1, :valign => 'top' )
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
    worksheet1.set_column( 'C:C', 25 )
    worksheet1.set_row( 2, 50 )
    worksheet1.set_row( 5, 50 )


    # Simple ascii string.
    cell_text = 'Hold the mouse over this cell to see the comment.'

    comment = 'This is a comment.'

    worksheet1.write( 'C3', cell_text, text_wrap )
    worksheet1.write_comment( 'C3', comment )

    cell_text = 'This is a UTF-8 string.'
    comment   = '☺'

    worksheet1.write( 'C6', cell_text, text_wrap )
    worksheet1.write_comment( 'C6', comment )



    ###############################################################################
    #
    # Example 2. Demonstrates visible and hidden comments.
    #

    # Set up some formatting.
    worksheet2.set_column( 'C:C', 25 )
    worksheet2.set_row( 2, 50 )
    worksheet2.set_row( 5, 50 )


    cell_text = 'This cell comment is visible.'

    comment = 'Hello.'

    worksheet2.write( 'C3', cell_text, text_wrap )
    worksheet2.write_comment( 'C3', comment, :visible => 1 )


    cell_text = "This cell comment isn't visible (the default)."

    comment = 'Hello.'

    worksheet2.write( 'C6', cell_text, text_wrap )
    worksheet2.write_comment( 'C6', comment )


    ###############################################################################
    #
    # Example 3. Demonstrates visible and hidden comments set at the worksheet
    #            level.
    #

    # Set up some formatting.
    worksheet3.set_column( 'C:C', 25 )
    worksheet3.set_row( 2, 50 )
    worksheet3.set_row( 5, 50 )
    worksheet3.set_row( 8, 50 )

    # Make all comments on the worksheet visible.
    worksheet3.show_comments

    cell_text = 'This cell comment is visible, explicitly.'

    comment = 'Hello.'

    worksheet3.write( 'C3', cell_text, text_wrap )
    worksheet3.write_comment( 'C3', comment, :visible => 1 )


    cell_text =
      'This cell comment is also visible because we used show_comments().'

    comment = 'Hello.'

    worksheet3.write( 'C6', cell_text, text_wrap )
    worksheet3.write_comment( 'C6', comment )


    cell_text = 'However, we can still override it locally.'

    comment = 'Hello.'

    worksheet3.write( 'C9', cell_text, text_wrap )
    worksheet3.write_comment( 'C9', comment, :visible => 0 )


    ###############################################################################
    #
    # Example 4. Demonstrates changes to the comment box dimensions.
    #

    # Set up some formatting.
    worksheet4.set_column( 'C:C', 25 )
    worksheet4.set_row( 2,  50 )
    worksheet4.set_row( 5,  50 )
    worksheet4.set_row( 8,  50 )
    worksheet4.set_row( 15, 50 )

    worksheet4.show_comments

    cell_text = 'This cell comment is default size.'

    comment = 'Hello.'

    worksheet4.write( 'C3', cell_text, text_wrap )
    worksheet4.write_comment( 'C3', comment )


    cell_text = 'This cell comment is twice as wide.'

    comment = 'Hello.'

    worksheet4.write( 'C6', cell_text, text_wrap )
    worksheet4.write_comment( 'C6', comment, :x_scale => 2 )


    cell_text = 'This cell comment is twice as high.'

    comment = 'Hello.'

    worksheet4.write( 'C9', cell_text, text_wrap )
    worksheet4.write_comment( 'C9', comment, :y_scale => 2 )


    cell_text = 'This cell comment is scaled in both directions.'

    comment = 'Hello.'

    worksheet4.write( 'C16', cell_text, text_wrap )
    worksheet4.write_comment( 'C16', comment, :x_scale => 1.2, :y_scale => 0.8 )


    cell_text = 'This cell comment has width and height specified in pixels.'

    comment = 'Hello.'

    worksheet4.write( 'C19', cell_text, text_wrap )
    worksheet4.write_comment( 'C19', comment, :width => 200, :height => 20 )


    ###############################################################################
    #
    # Example 5. Demonstrates changes to the cell comment position.
    #

    worksheet5.set_column( 'C:C', 25 )
    worksheet5.set_row( 2,  50 )
    worksheet5.set_row( 5,  50 )
    worksheet5.set_row( 8,  50 )
    worksheet5.set_row( 11, 50 )

    worksheet5.show_comments

    cell_text = 'This cell comment is in the default position.'

    comment = 'Hello.'

    worksheet5.write( 'C3', cell_text, text_wrap )
    worksheet5.write_comment( 'C3', comment )


    cell_text = 'This cell comment has been moved to another cell.'

    comment = 'Hello.'

    worksheet5.write( 'C6', cell_text, text_wrap )
    worksheet5.write_comment( 'C6', comment, :start_cell => 'E4' )


    cell_text = 'This cell comment has been moved to another cell.'

    comment = 'Hello.'

    worksheet5.write( 'C9', cell_text, text_wrap )
    worksheet5.write_comment( 'C9', comment, :start_row => 8, :start_col => 4 )


    cell_text = 'This cell comment has been shifted within its default cell.'

    comment = 'Hello.'

    worksheet5.write( 'C12', cell_text, text_wrap )
    worksheet5.write_comment( 'C12', comment, :x_offset => 30, :y_offset => 12 )


    ###############################################################################
    #
    # Example 6. Demonstrates changes to the comment background colour.
    #

    worksheet6.set_column( 'C:C', 25 )
    worksheet6.set_row( 2, 50 )
    worksheet6.set_row( 5, 50 )
    worksheet6.set_row( 8, 50 )

    worksheet6.show_comments

    cell_text = 'This cell comment has a different colour.'

    comment = 'Hello.'

    worksheet6.write( 'C3', cell_text, text_wrap )
    worksheet6.write_comment( 'C3', comment, :color => 'green' )


    cell_text = 'This cell comment has the default colour.'

    comment = 'Hello.'

    worksheet6.write( 'C6', cell_text, text_wrap )
    worksheet6.write_comment( 'C6', comment )


    cell_text = 'This cell comment has a different colour.'

    comment = 'Hello.'

    worksheet6.write( 'C9', cell_text, text_wrap )
    worksheet6.write_comment( 'C9', comment, :color => 0x35 )


    ###############################################################################
    #
    # Example 7. Demonstrates how to set the cell comment author.
    #

    worksheet7.set_column( 'C:C', 30 )
    worksheet7.set_row( 2,  50 )
    worksheet7.set_row( 5,  50 )
    worksheet7.set_row( 8,  50 )

    author = ''
    cell   = 'C3'

    cell_text = "Move the mouse over this cell and you will see 'Cell commented " +
      "by #{author}' (blank) in the status bar at the bottom"

    comment = 'Hello.'

    worksheet7.write( cell, cell_text, text_wrap )
    worksheet7.write_comment( cell, comment )


    author    = 'Ruby'
    cell      = 'C6'
    cell_text = "Move the mouse over this cell and you will see 'Cell commented " +
      "by #{author}' in the status bar at the bottom"

    comment = 'Hello.'

    worksheet7.write( cell, cell_text, text_wrap )
    worksheet7.write_comment( cell, comment, :author => author )


    author    = '€'
    cell      = 'C9'
    cell_text = "Move the mouse over this cell and you will see 'Cell commented " +
      "by #{author}' in the status bar at the bottom"
    comment = 'Hello.'

    worksheet7.write( cell, cell_text, text_wrap )
    worksheet7.write_comment( cell, comment, :author => author )




    ###############################################################################
    #
    # Example 8. Demonstrates the need to explicitly set the row height.
    #

    # Set up some formatting.
    worksheet8.set_column( 'C:C', 25 )
    worksheet8.set_row( 2, 80 )

    worksheet8.show_comments


    cell_text =
      'The height of this row has been adjusted explicitly using ' +
      'set_row(). The size of the comment box is adjusted ' +
      'accordingly by WriteXLSX.'

    comment = 'Hello.'

    worksheet8.write( 'C3', cell_text, text_wrap )
    worksheet8.write_comment( 'C3', comment )


    cell_text =
      'The height of this row has been adjusted by Excel due to the ' +
      'text wrap property being set. Unfortunately this means that ' +
      'the height of the row is unknown to WriteXLSX at ' +
      "run time and thus the comment box is stretched as well.\n\n" +
      'Use set_row() to specify the row height explicitly to avoid ' +
      'this problem.'

    comment = 'Hello.'

    worksheet8.write( 'C6', cell_text, text_wrap )
    worksheet8.write_comment( 'C6', comment )

    workbook.close
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_rich_strings
    xlsx = 'rich_strings.xlsx'
    workbook  = WriteXLSX.new(xlsx)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
  end

  def test_autofilter
    xlsx = 'autofilter.xlsx'
    workbook = WriteXLSX.new(xlsx)

    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet

    bold = workbook.add_format( :bold => 1 )

    # Extract the data embedded at the end of this file.
    data_array = autofilter_data.split("\n")
    headings = data_array.shift.split
    data = []
    data_array.each { |line| data << line.split }

    # Set up several sheets with the same data.
    workbook.worksheets.each do |worksheet|
      worksheet.set_column( 'A:D', 12 )
      worksheet.set_row( 0, 20, bold )
      worksheet.write( 'A1', headings )
    end


    ###############################################################################
    #
    # Example 1. Autofilter without conditions.
    #

    worksheet1.autofilter( 'A1:D51' )
    worksheet1.write( 'A2', [ data ] )

    ###############################################################################
    #
    #
    # Example 2. Autofilter with a filter condition in the first column.
    #

    # The range in this example is the same as above but in row-column notation.
    worksheet2.autofilter( 0, 0, 50, 3 )

    # The placeholder "Region" in the filter is ignored and can be any string
    # that adds clarity to the expression.
    #
    worksheet2.filter_column( 0, 'Region eq East' )

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

    worksheet3.autofilter( 'A1:D51' )

    worksheet3.filter_column( 'A', 'x eq East or x eq South' )

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
      region = row_data[0]

      worksheet3.set_row(row, nil, nil, 1) unless region == 'East' || region == 'South'
      worksheet3.write(row, 0,row_data)
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
    compare_xlsx(@expected_dir, @result_dir, xlsx)
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

  def compare_xlsx(expected, result, xlsx)
    begin
      prepare_compare(expected, result, xlsx)
      expected_files = files(expected)
      result_files   = files(result)

      not_exists = expected_files - result_files
      assert(not_exists.empty?, "These files does not exist: #{not_exists.to_s}")

      additional_exist = result_files - expected_files
      assert(additional_exist.empty?, "These files must not exist: #{additional_exist.to_s}")

      expected_files.each do |file|
        assert_equal(got_to_array(IO.read(File.join(expected, file))),
                     got_to_array(IO.read(File.join(result, file))),
                     "#{file} differs.")
      end
    ensure
      cleanup(xlsx)
    end
  end

  def prepare_compare(expected, result, xlsx)
    prepare_xlsx(expected, File.join(@perl_output, xlsx))
    prepare_xlsx(result, xlsx)
  end

  def prepare_xlsx(dir, xlsx)
    Dir.mkdir(dir)
    system("unzip -q #{xlsx} -d #{dir}")
  end

  def files(dir)
    Dir.glob(File.join(dir, "**/*")).select { |f| File.file?(f) }.
                                     reject { |f| File.basename(f) =~ /(core|theme1)\.xml/ }.
                                     collect { |f| f.sub(Regexp.new("^#{dir}"), '') }
  end

  def cleanup(xlsx)
    system("rm -rf #{xlsx}")          if File.exist?(xlsx)
    system("rm -rf #{@expected_dir}") if File.exist?(@expected_dir)
    system("rm -rf #{@result_dir}")   if File.exist?(@result_dir)
  end
end
