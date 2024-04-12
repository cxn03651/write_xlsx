# write_xlsx

[![Gem Version](https://badge.fury.io/rb/write_xlsx.png)](http://badge.fury.io/rb/write_xlsx)
[![Build Status](https://travis-ci.org/cxn03651/write_xlsx.svg?branch=master)](https://travis-ci.org/cxn03651/write_xlsx)

gem to create a new file in the Excel 2007+ XLSX format, and you can use the
same interface as writeexcel gem. write_xlsx is converted from Perl's module
[Excel::Writer::XLSX](https://github.com/jmcnamara/excel-writer-xlsx)

## Description

Reference doc : https://cxn03651.github.io/write_xlsx/

The WriteXLSX supports the following features:
* Multiple worksheets
* Strings and numbers
* Unicode text
* Cell formatting
* Formulas (including array formats)
* Images
* Charts
* Autofilters
* Data validation
* Conditional formatting
* Macros
* Tables
* Shapes
* Sparklines
* Hyperlinks
* Rich string formats
* Defined names
* Grouping/Outlines
* Cell comments
* Panes
* Page set-up and printing options

write_xlsx uses the same interface as writeexcel gem.

## Installation

Add this line to your application's Gemfile:

    gem 'write_xlsx'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install write_xlsx

## Synopsis

To write a string, a formatted string, a number and a formula to
the first worksheet in an Excel XML spreadsheet called ruby.xlsx:

    require 'rubygems'
    require 'write_xlsx'

    # Create a new Excel workbook
    workbook = WriteXLSX.new('ruby.xlsx')

    # Add a worksheet
    worksheet = workbook.add_worksheet

    # Add and define a format
    format = workbook.add_format # Add a format
    format.set_bold
    format.set_color('red')
    format.set_align('center')

    # Write a formatted and unformatted string, row and column notation.
    col = row = 0
    worksheet.write(row, col, "Hi Excel!", format)
    worksheet.write(1,   col, "Hi Excel!")

    # Write a number and a formula using A1 notation
    worksheet.write('A3', 1.2345)
    worksheet.write('A4', '=SIN(PI()/4)')

    workbook.close

## Copyright
Original Perl module was written by John McNamara(jmcnamara@cpan.org).

Converted to ruby by Hideo NAKAMURA(nakamrua.hideo@gmail.com)
Copyright (c) 2012-2024 Hideo NAKAMURA.

See LICENSE.txt for further details.

## Contributing to write_xlsx

* repsitory: https://github.com/cxn03651/write_xlsx
* Check out the latest master to make sure the feature hasn't been implemented or the bug hasn't been fixed yet
* Check out the issue tracker to make sure someone already hasn't requested it and/or contributed it
* Fork the project
* Start a feature/bugfix branch
* Commit and push until you are happy with your contribution
* Make sure to add tests for it. This is important so I don't break it in a future version unintentionally.
* Please try not to mess with the Rakefile, version, or history. If you want to have your own version, or is otherwise necessary, that is fine, but please isolate to its own commit so I can cherry-pick around it.
