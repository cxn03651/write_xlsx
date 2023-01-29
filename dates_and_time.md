---
layout: default
title: Dates and Time
---
#### <a name="dates_and_time" class="anchor" href="#dates_and_time"><span class="octicon octicon-link" /></a>DATES AND TIME IN EXCEL

There are two important things to understand about dates and times in Excel:

1. A date/time in Excel is a real number plus an Excel number format.
2. WriteXLSX doesn't automatically convert date/time strings in `write()`
   to an Excel date/time.

These two points are explained in more detail below along with some suggestions
on how to convert times and dates to the required format.

##### An Excel date/time is a number plus a format

If you write a date string with `write()` then all you will get is a string:

    worksheet.write('A1', '02/03/04')   # !! Writes a string not a date. !!

Dates and times in Excel are represented by real numbers,
for example "Jan 1 2001 12:30 AM" is represented by the number 36892.521.

The integer part of the number stores the number of days since the epoch
and the fractional part stores the percentage of the day.

A date or time in Excel is just like any other number.
To have the number display as a date you must apply an Excel number format to it.
Here are some examples.

    require 'write_xlsx'

    workbook  = WriteXLSX.new('date_examples.xlsx')
    worksheet = workbook.add_worksheet

    worksheet.set_column('A:A', 30)    # For extra visibility.

    number = 39506.5;

    worksheet.write('A1', number)             #   39506.5

    format2 = workbook.add_format(num_format: 'dd/mm/yy')
    worksheet.write('A2', number, format2)    #  28/02/08

    format3 = workbook.add_format(num_format: 'mm/dd/yy')
    worksheet.write('A3', number, format3)    #  02/28/08

    format4 = workbook.add_format(num_format: 'd-m-yyyy')
    worksheet.write('A4', number, format4)    #  28-2-2008

    format5 = workbook.add_format(num_format: 'dd/mm/yy hh:mm')
    worksheet.write('A5', number, format5)    #  28/02/08 12:00

    format6 = workbook.add_format(num_format: 'd mmm yyyy')
    worksheet.write('A6', number, format6)    # 28 Feb 2008

    format7 = workbook.add_format(num_format: 'mmm d yyyy hh:mm AM/PM')
    worksheet.write('A7', number , format7);     #  Feb 28 2008 12:00 PM

##### WriteXLSX doesn't automatically convert date/time strings

WriteXLSX doesn't automatically convert input date strings into Excel's
formatted date numbers due to the large number of possible date formats
and also due to the possibility of misinterpretation.

For example, does 02/03/04 mean March 2 2004, February 3 2004
or even March 4 2002.

Therefore, in order to handle dates you will have to convert them to numbers
and apply an Excel format.
Some methods for converting dates are listed in the next section.

The most direct way is to convert your dates to the
ISO8601 yyyy-mm-ddThh:mm:ss.sss date format and use the
`write_date_time()` worksheet method:

    worksheet.write_date_time('A2', '2001-01-01T12:20', format)

See the `write_date_time()` section of the documentation for more details.

A general methodology for handling date strings with `write_date_time()` is:

1. Identify incoming date/time strings with a regex.
2. Extract the component parts of the date/time using the same regex.
3. Convert the date/time to the ISO8601 format.
4. Write the date/time using `write_date_time()` and a number format.

Here is an example:

    require 'write_xlsx'

    workbook  = WriteXLSX.new('example.xlsx')
    worksheet = workbook.add_worksheet

    # Set the default format for dates.
    date_format = workbook.add_format(num_format: 'mmm d yyyy')

    # Increase column width to improve visibility of data.
    worksheet.set_column('A:C', 20)

    row = 0
    while(line = gets)
      col  = 0;
      data = line.chomp.split

      data.each do |item|
        # Match dates in the following formats: d/m/yy, d/m/yyyy
        if item =~ %r{\A(\d{1,2})/(\d{1,2})/(\d{4})\Z}
          # Change to the date format required by write_date_time().
          date_str = sprintf("%4d-%02d-%02dT", $3, $2, $1)
          worksheet.write_date_time(row, col, date, date_format)
        else
          # Just plain data
          worksheet.write(row, col, item)
        end
        col += 1
      end
      row += 1
    end

    ##-- data file --
    Item    Cost    Date
    Book    10      1/9/2007
    Beer    4       12/9/2007
    Bed     500     5/10/2007

##### Converting dates and times to an Excel date or time

The `write_date_time()` method above is just one way of handling dates and times.

You can also use the `convert_date_time()` worksheet method to convert from
an ISO8601 style date string to an Excel date and time number.
