# -*- coding: utf-8 -*-
require 'helper'

class TestUrlsAsStrings < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_urls_as_strings
    @xlsx = 'urls_as_strings.xlsx'
    workbook = WriteXLSX.new(@xlsx, urls_as_strings: true)
    worksheet = workbook.add_worksheet
    worksheet.write('A1', 'http://www.write_xlsx.com')
    worksheet.write('A2', 'mailto:write_xlsx@example.com')
    worksheet.write('A3', 'ftp://ftp.ruby.org/' )
    worksheet.write('A4', 'internal:Sheet1!A1'  )
    worksheet.write('A5', 'external:c:\foo.xlsx')
    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
