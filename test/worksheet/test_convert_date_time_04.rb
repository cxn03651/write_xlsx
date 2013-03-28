# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestConvertDateTime04 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_convert_date_time_should_not_change_date_time_string
    date_time = ' 2000-01-23T00:00:00.000Z '
    @worksheet.convert_date_time(date_time)

    assert_equal(' 2000-01-23T00:00:00.000Z ', date_time)
  end
end
