# -*- coding: utf-8 -*-
require 'helper'

class TestUrlsAsStrings < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_urls_as_strings
    @xlsx = 'urls_as_strings.xlsx'
    workbook = WriteXLSX.new(@io, strings_to_urls: false)
    worksheet = workbook.add_worksheet
    worksheet.write('A1', 'http://www.write_xlsx.com')
    worksheet.write('A2', 'mailto:write_xlsx@example.com')
    worksheet.write('A3', 'ftp://ftp.ruby.org/' )
    worksheet.write('A4', 'internal:Sheet1!A1'  )
    worksheet.write('A5', 'external:c:\foo.xlsx')
    workbook.close
    compare_for_regression
  end
end
