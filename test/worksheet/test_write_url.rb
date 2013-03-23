# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteUrl < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_url_does_not_change_url
    url = 'external:c:\temp\foo.xlsx#my_name'.freeze
    assert_nothing_raised do
      @worksheet.write_url('A1', url)
    end
  end
end
